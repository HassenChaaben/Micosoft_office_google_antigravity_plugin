import os
import sys
import subprocess
from unittest.mock import patch
import pytest
from cryptography import x509
from cryptography.hazmat.primitives.asymmetric import padding
from cryptography.x509.oid import ExtensionOID, ExtendedKeyUsageOID, NameOID

from daemon.cert_gen import (
    generate_ca_cert,
    generate_localhost_cert,
    install_ca_to_windows_store,
    install_localhost_to_windows_store,
    install_cert_to_windows_store,
    setup_certs,
)


def test_generate_ca_cert(tmp_path):
    cert_dir = str(tmp_path / "certs")
    ca_crt, ca_key = generate_ca_cert(cert_dir=cert_dir)

    assert os.path.exists(ca_crt)
    assert os.path.exists(ca_key)

    with open(ca_crt, "rb") as f:
        cert = x509.load_pem_x509_certificate(f.read())

    # Verify self-signed root CA properties
    assert cert.subject == cert.issuer
    cn = cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value
    assert cn == "Gemini Antigravity Local CA"

    bc = cert.extensions.get_extension_for_oid(ExtensionOID.BASIC_CONSTRAINTS).value
    assert bc.ca is True

    ku = cert.extensions.get_extension_for_oid(ExtensionOID.KEY_USAGE).value
    assert ku.key_cert_sign is True
    assert ku.crl_sign is True

    ski = cert.extensions.get_extension_for_oid(ExtensionOID.SUBJECT_KEY_IDENTIFIER).value
    assert ski is not None


def test_generate_localhost_cert_signed_by_ca(tmp_path):
    cert_dir = str(tmp_path / "certs")
    loc_crt, loc_key = generate_localhost_cert(cert_dir=cert_dir)

    assert os.path.exists(loc_crt)
    assert os.path.exists(loc_key)

    ca_crt = os.path.join(cert_dir, "ca.crt")
    assert os.path.exists(ca_crt)

    with open(loc_crt, "rb") as f:
        loc_cert = x509.load_pem_x509_certificate(f.read())

    with open(ca_crt, "rb") as f:
        ca_cert = x509.load_pem_x509_certificate(f.read())

    # Verify localhost certificate is issued and signed by Root CA
    assert loc_cert.issuer == ca_cert.subject
    cn = loc_cert.subject.get_attributes_for_oid(NameOID.COMMON_NAME)[0].value
    assert cn == "localhost"

    # Verify signature
    ca_public_key = ca_cert.public_key()
    ca_public_key.verify(
        loc_cert.signature,
        loc_cert.tbs_certificate_bytes,
        padding.PKCS1v15(),
        loc_cert.signature_hash_algorithm,
    )

    # Verify constraints & usage
    bc = loc_cert.extensions.get_extension_for_oid(ExtensionOID.BASIC_CONSTRAINTS).value
    assert bc.ca is False

    ku = loc_cert.extensions.get_extension_for_oid(ExtensionOID.KEY_USAGE).value
    assert ku.digital_signature is True
    assert ku.key_encipherment is True
    assert ku.key_cert_sign is False

    eku = loc_cert.extensions.get_extension_for_oid(ExtensionOID.EXTENDED_KEY_USAGE).value
    assert ExtendedKeyUsageOID.SERVER_AUTH in eku

    # Verify Subject Alternative Name
    san = loc_cert.extensions.get_extension_for_oid(ExtensionOID.SUBJECT_ALTERNATIVE_NAME).value
    dns_names = san.get_values_for_type(x509.DNSName)
    ip_addrs = [str(ip) for ip in san.get_values_for_type(x509.IPAddress)]

    assert "localhost" in dns_names
    assert "127.0.0.1" in ip_addrs
    assert "::1" in ip_addrs


def test_reuse_existing_certs_unless_force(tmp_path):
    cert_dir = str(tmp_path / "certs")
    ca_crt1, _ = generate_ca_cert(cert_dir=cert_dir)
    loc_crt1, _ = generate_localhost_cert(cert_dir=cert_dir)

    with open(loc_crt1, "rb") as f:
        serial1 = x509.load_pem_x509_certificate(f.read()).serial_number

    # Without force: should keep existing serial
    loc_crt2, _ = generate_localhost_cert(cert_dir=cert_dir, force=False)
    with open(loc_crt2, "rb") as f:
        serial2 = x509.load_pem_x509_certificate(f.read()).serial_number
    assert serial1 == serial2

    # With force: should regenerate with fresh serial
    loc_crt3, _ = generate_localhost_cert(cert_dir=cert_dir, force=True)
    with open(loc_crt3, "rb") as f:
        serial3 = x509.load_pem_x509_certificate(f.read()).serial_number
    assert serial1 != serial3


def test_missing_ca_key_triggers_regeneration(tmp_path):
    cert_dir = str(tmp_path / "certs")
    generate_ca_cert(cert_dir=cert_dir)

    # Delete ca.key to simulate shipped cert without key
    os.remove(os.path.join(cert_dir, "ca.key"))
    assert not os.path.exists(os.path.join(cert_dir, "ca.key"))

    # Generating localhost cert should regenerate CA and succeed
    loc_crt, loc_key = generate_localhost_cert(cert_dir=cert_dir)
    assert os.path.exists(loc_crt)
    assert os.path.exists(loc_key)
    assert os.path.exists(os.path.join(cert_dir, "ca.key"))


def test_install_ca_and_localhost_to_windows_store(tmp_path):
    cert_dir = str(tmp_path / "certs")
    ca_crt, ca_key = generate_ca_cert(cert_dir=cert_dir)
    loc_crt, loc_key = generate_localhost_cert(cert_dir=cert_dir, ca_cert_path=ca_crt, ca_key_path=ca_key)

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0

        # Test CA installation into CurrentUser\Root
        assert install_ca_to_windows_store(ca_crt) is True

        user_root_called = any(
            call.args[0] == ["certutil", "-addstore", "-f", "-user", "Root", ca_crt]
            for call in mock_run.call_args_list
        )
        assert user_root_called

        # Test localhost installation into CurrentUser\TrustedPeople
        assert install_localhost_to_windows_store(loc_crt) is True
        trusted_people_called = any(
            call.args[0] == ["certutil", "-addstore", "-f", "-user", "TrustedPeople", loc_crt]
            for call in mock_run.call_args_list
        )
        assert trusted_people_called


def test_setup_certs_turnkey(tmp_path):
    cert_dir = str(tmp_path / "certs")

    with patch("subprocess.run") as mock_run:
        mock_run.return_value.returncode = 0
        ca_crt, loc_crt, loc_key = setup_certs(cert_dir=cert_dir)

        assert os.path.exists(ca_crt)
        assert os.path.exists(loc_crt)
        assert os.path.exists(loc_key)

        # Both user Root and TrustedPeople stores were targeted
        calls = [c.args[0] for c in mock_run.call_args_list]
        assert any("-user" in cmd and "Root" in cmd for cmd in calls)
        assert any("-user" in cmd and "TrustedPeople" in cmd for cmd in calls)


def test_install_missing_file_returns_false(tmp_path):
    missing_path = str(tmp_path / "non_existent.crt")
    assert install_ca_to_windows_store(missing_path) is False
    assert install_localhost_to_windows_store(missing_path) is False
    assert install_cert_to_windows_store(cert_path=missing_path, ca_path=missing_path) is False


def test_cli_setup_cert_flag():
    # Verify CLI execution with --setup-cert flag exits cleanly (0) with mocked subprocess
    script = (
        "import sys, unittest.mock as mock\n"
        "with mock.patch('subprocess.run') as m:\n"
        "    m.return_value.returncode = 0\n"
        "    import runpy\n"
        "    sys.argv = ['daemon.cert_gen', '--setup-cert']\n"
        "    runpy.run_module('daemon.cert_gen', run_name='__main__')\n"
    )
    cmd = [sys.executable, "-c", script]
    res = subprocess.run(cmd, capture_output=True, text=True)
    assert res.returncode == 0
    assert "[CertGen]" in res.stdout
