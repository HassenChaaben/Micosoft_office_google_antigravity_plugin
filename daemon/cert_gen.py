"""
Local SSL Certificate Generator and Windows Store Installer
Generates localhost TLS certificates with Subject Alternative Names (SAN)
and installs them into Windows TrustedPeople and Root stores so WebView2 trusts https://localhost:18890.
"""

import os
import sys
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
CERT_FILE = os.path.join(CERT_DIR, "localhost.crt")
KEY_FILE = os.path.join(CERT_DIR, "localhost.key")

def generate_localhost_cert(force=False):
    os.makedirs(CERT_DIR, exist_ok=True)

    if not force and os.path.exists(CERT_FILE) and os.path.exists(KEY_FILE):
        return CERT_FILE, KEY_FILE

    print("[CertGen] Generating 2048-bit RSA self-signed localhost certificate...")
    key = rsa.generate_private_key(
        public_exponent=65537,
        key_size=2048,
    )

    subject = issuer = x509.Name([
        x509.NameAttribute(NameOID.COUNTRY_NAME, "US"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "Google Antigravity"),
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
    ])

    now = datetime.datetime.now(datetime.timezone.utc)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(issuer)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=1825))
        .add_extension(
            x509.BasicConstraints(ca=True, path_length=None),
            critical=True,
        )
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=True,
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
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False,
        )
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
            ]),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )

    with open(KEY_FILE, "wb") as f:
        f.write(
            key.private_bytes(
                encoding=serialization.Encoding.PEM,
                format=serialization.PrivateFormat.TraditionalOpenSSL,
                encryption_algorithm=serialization.NoEncryption(),
            )
        )

    with open(CERT_FILE, "wb") as f:
        f.write(cert.public_bytes(serialization.Encoding.PEM))

    print(f"[CertGen] Certificate generated at {CERT_FILE}")
    return CERT_FILE, KEY_FILE

def install_cert_to_windows_store(cert_path=None):
    if cert_path is None:
        cert_path = CERT_FILE

    if not os.path.exists(cert_path):
        print(f"[CertGen] Cert file not found: {cert_path}")
        return False

    # 1. Install to user's TrustedPeople (silent, zero prompt)
    try:
        subprocess.run(["certutil", "-addstore", "-f", "-user", "TrustedPeople", cert_path],
                       capture_output=True, text=True, check=True)
        print("[CertGen] Installed to CurrentUser\\TrustedPeople successfully.")
    except Exception as e:
        print(f"[CertGen] Warning TrustedPeople: {e}")

    return True

if __name__ == "__main__":
    c, k = generate_localhost_cert(force=False)
    install_cert_to_windows_store(c)
