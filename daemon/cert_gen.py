"""
Local SSL Certificate Generator and Windows Store Installer
Generates per-machine CA and localhost TLS certificates with Subject Alternative Names (SAN)
and installs them into Windows CurrentUser\\Root and TrustedPeople stores so WebView2 trusts https://localhost:18890.
"""

import os
import sys
import argparse
import datetime
import ipaddress
import subprocess
from cryptography import x509
from cryptography.x509.oid import NameOID, ExtendedKeyUsageOID
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.hazmat.primitives import serialization

def get_base_dir() -> str:
    if getattr(sys, "frozen", False):
        exe_dir = os.path.dirname(sys.executable)
        if os.path.exists(os.path.join(exe_dir, "certs")):
            return exe_dir
        parent = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
        if os.path.exists(os.path.join(parent, "certs")):
            return parent
        return exe_dir
    return os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))

BASE_DIR = get_base_dir()
CERT_DIR = os.path.join(BASE_DIR, "certs")
CA_CERT_FILE = os.path.join(CERT_DIR, "ca.crt")
CA_KEY_FILE = os.path.join(CERT_DIR, "ca.key")
CERT_FILE = os.path.join(CERT_DIR, "localhost.crt")
KEY_FILE = os.path.join(CERT_DIR, "localhost.key")

def generate_ca_cert(cert_dir: str = None, force: bool = False):
    """
    Generate a local per-machine Root CA certificate and private key.
    """
    target_dir = cert_dir if cert_dir is not None else CERT_DIR
    os.makedirs(target_dir, exist_ok=True)

    ca_crt_path = os.path.join(target_dir, "ca.crt")
    ca_key_path = os.path.join(target_dir, "ca.key")

    if not force and os.path.exists(ca_crt_path) and os.path.exists(ca_key_path):
        return ca_crt_path, ca_key_path

    print("[CertGen] Generating 2048-bit RSA per-machine Root CA...")
    ca_key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    ca_name = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Google Antigravity"),
        x509.NameAttribute(NameOID.COMMON_NAME, "Gemini Antigravity Local CA"),
    ])

    now = datetime.datetime.now(datetime.timezone.utc)
    ca_cert = (
        x509.CertificateBuilder()
        .subject_name(ca_name)
        .issuer_name(ca_name)
        .public_key(ca_key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=3650))
        .add_extension(
            x509.BasicConstraints(ca=True, path_length=0),
            critical=True,
        )
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=False,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=True,
                crl_sign=True,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(ca_key.public_key()),
            critical=False,
        )
        .sign(ca_key, hashes.SHA256())
    )

    with open(ca_key_path, "wb") as f:
        f.write(
            ca_key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption(),
            )
        )

    with open(ca_crt_path, "wb") as f:
        f.write(ca_cert.public_bytes(serialization.Encoding.PEM))

    print(f"[CertGen] Per-machine Root CA generated at {ca_crt_path}")
    return ca_crt_path, ca_key_path

def generate_localhost_cert(force: bool = False, cert_dir: str = None, ca_cert_path: str = None, ca_key_path: str = None):
    """
    Generate localhost TLS certificate signed by the per-machine Root CA.
    """
    target_dir = cert_dir if cert_dir is not None else CERT_DIR
    os.makedirs(target_dir, exist_ok=True)

    loc_crt_path = os.path.join(target_dir, "localhost.crt")
    loc_key_path = os.path.join(target_dir, "localhost.key")

    ca_crt = ca_cert_path if ca_cert_path is not None else os.path.join(target_dir, "ca.crt")
    ca_key = ca_key_path if ca_key_path is not None else os.path.join(target_dir, "ca.key")

    # If CA key or CA cert missing, regenerate CA
    if not os.path.exists(ca_crt) or not os.path.exists(ca_key):
        ca_crt, ca_key = generate_ca_cert(cert_dir=target_dir, force=force)

    if not force and os.path.exists(loc_crt_path) and os.path.exists(loc_key_path):
        return loc_crt_path, loc_key_path

    print("[CertGen] Generating 2048-bit RSA localhost certificate signed by Root CA...")
    with open(ca_key, "rb") as f:
        ca_private_key = serialization.load_pem_private_key(f.read(), password=None)

    with open(ca_crt, "rb") as f:
        ca_cert_obj = x509.load_pem_x509_certificate(f.read())

    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    subject = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Google Antigravity"),
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])

    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(ca_cert_obj.subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=1825))
        .add_extension(
            x509.BasicConstraints(ca=False, path_length=None),
            critical=True,
        )
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=True,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False,
        )
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
                x509.IPAddress(ipaddress.IPv6Address("::1")),
            ]),
            critical=False,
        )
        .add_extension(
            x509.AuthorityKeyIdentifier.from_issuer_public_key(ca_private_key.public_key()),
            critical=False,
        )
        .add_extension(
            x509.SubjectKeyIdentifier.from_public_key(key.public_key()),
            critical=False,
        )
        .sign(ca_private_key, hashes.SHA256())
    )

    with open(loc_key_path, "wb") as f:
        f.write(
            key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption(),
            )
        )

    with open(loc_crt_path, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    print(f"[CertGen] Certificate generated at {loc_crt_path}")
    return loc_crt_path, loc_key_path

def install_ca_to_windows_store(ca_path: str = None) -> bool:
    """
    Install Root CA certificate into Windows CurrentUser\\Root store.
    Also attempts LocalMachine\\Root store if running elevated.
    """
    if ca_path is None:
        ca_path = CA_CERT_FILE

    if not os.path.exists(ca_path):
        print(f"[CertGen] CA cert file not found: {ca_path}")
        return False

    success = False
    # 1. Install to CurrentUser\Root (primary for user profile / WebView2)
    try:
        subprocess.run(["certutil", "-addstore", "-f", "-user", "Root", ca_path],
                       capture_output=True, text=True, check=True, timeout=5)
        print("[CertGen] Installed CA to CurrentUser\\Root successfully.")
        success = True
    except Exception as e:
        print(f"[CertGen] Warning CurrentUser\\Root: {e}")

    # 2. Attempt LocalMachine\Root (useful during admin installer setup)
    try:
        subprocess.run(["certutil", "-addstore", "-f", "Root", ca_path],
                       capture_output=True, text=True, check=True, timeout=5)
        print("[CertGen] Installed CA to LocalMachine\\Root successfully.")
        success = True
    except Exception:
        pass

    return success

def install_localhost_to_windows_store(cert_path: str = None) -> bool:
    """
    Install localhost leaf certificate into Windows CurrentUser\\TrustedPeople store.
    """
    if cert_path is None:
        cert_path = CERT_FILE

    if not os.path.exists(cert_path):
        print(f"[CertGen] Cert file not found: {cert_path}")
        return False

    try:
        subprocess.run(["certutil", "-addstore", "-f", "-user", "TrustedPeople", cert_path],
                       capture_output=True, text=True, check=True, timeout=5)
        print("[CertGen] Installed localhost cert to CurrentUser\\TrustedPeople successfully.")
        return True
    except Exception as e:
        print(f"[CertGen] Warning TrustedPeople: {e}")
        return False

def install_cert_to_windows_store(cert_path: str = None, ca_path: str = None) -> bool:
    """
    Installs CA into CurrentUser\\Root and localhost cert into CurrentUser\\TrustedPeople.
    Maintains full backward compatibility with single cert_path argument.
    """
    if ca_path is None:
        ca_path = CA_CERT_FILE
    if cert_path is None:
        cert_path = CERT_FILE

    ca_ok = False
    if os.path.exists(ca_path):
        ca_ok = install_ca_to_windows_store(ca_path)

    loc_ok = False
    if os.path.exists(cert_path):
        loc_ok = install_localhost_to_windows_store(cert_path)

    return ca_ok or loc_ok

def setup_certs(cert_dir: str = None, force: bool = False):
    """
    Turnkey installer entrypoint:
    Generates per-machine CA, generates localhost certificate signed by that CA,
    and installs CA into CurrentUser\\Root and localhost cert into CurrentUser\\TrustedPeople.
    """
    ca_crt, ca_key = generate_ca_cert(cert_dir=cert_dir, force=force)
    loc_crt, loc_key = generate_localhost_cert(force=force, cert_dir=cert_dir, ca_cert_path=ca_crt, ca_key_path=ca_key)
    install_cert_to_windows_store(cert_path=loc_crt, ca_path=ca_crt)
    print(f"[CertGen] Setup complete:\n  CA: {ca_crt}\n  Cert: {loc_crt}")
    return ca_crt, loc_crt, loc_key

# Early hook if daemon executable is invoked with --setup-cert directly
if "--setup-cert" in sys.argv and not any("pytest" in arg.lower() for arg in sys.argv):
    force_flag = "--force" in sys.argv
    setup_certs(force=force_flag)
    sys.exit(0)

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Gemini Antigravity Localhost TLS Certificate Generator")
    parser.add_argument("--setup-cert", action="store_true", help="Generate per-machine CA and localhost certificate and install into CurrentUser\\Root")
    parser.add_argument("--force", action="store_true", help="Force regenerate certificates")
    args, _ = parser.parse_known_args()

    setup_certs(force=args.force)
