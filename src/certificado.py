"""
Certificado Digital - Leitura de arquivos .pfx/.p12
Extrai CNPJ, titular, validade e gera arquivos PEM temporários para SSL.
"""

import re
import os
import tempfile
from pathlib import Path
from datetime import datetime, timezone

try:
    from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption
    from cryptography.hazmat.primitives.serialization import BestAvailableEncryption
    from cryptography import x509
    HAS_CRYPTOGRAPHY = True
except ImportError:
    HAS_CRYPTOGRAPHY = False

try:
    from OpenSSL import crypto as openssl_crypto
    HAS_OPENSSL = True
except ImportError:
    HAS_OPENSSL = False


class CertificadoDigital:
    """
    Carrega e manipula certificados digitais A1 (.pfx / .p12).
    Suporta as libs 'cryptography' e 'pyOpenSSL'.
    """

    def __init__(self, caminho: str, senha: str):
        self.caminho = caminho
        self.senha = senha
        self._senha_bytes = senha.encode("utf-8") if isinstance(senha, str) else senha

        self._cert_pem: bytes = b""
        self._key_pem: bytes = b""
        self._titular: str = ""
        self._cnpj: str = ""
        self._validade: str = ""

        self._carregar()

    def _carregar(self):
        with open(self.caminho, "rb") as f:
            pfx_data = f.read()

        if HAS_CRYPTOGRAPHY:
            self._carregar_cryptography(pfx_data)
        elif HAS_OPENSSL:
            self._carregar_pyopenssl(pfx_data)
        else:
            raise ImportError(
                "Instale 'cryptography' (pip install cryptography) "
                "ou 'pyOpenSSL' (pip install pyOpenSSL)."
            )

    def _carregar_cryptography(self, pfx_data: bytes):
        private_key, certificate, chain = pkcs12.load_key_and_certificates(
            pfx_data, self._senha_bytes
        )

        self._cert_pem = certificate.public_bytes(Encoding.PEM)
        self._key_pem = private_key.private_bytes(
            encoding=Encoding.PEM,
            format=PrivateFormat.PKCS8,
            encryption_algorithm=NoEncryption(),
        )

        # Titular (CN do Subject)
        try:
            cn = certificate.subject.get_attributes_for_oid(
                x509.NameOID.COMMON_NAME
            )
            self._titular = cn[0].value if cn else "Desconhecido"
        except Exception:
            self._titular = "Desconhecido"

        # Validade
        try:
            not_after = certificate.not_valid_after_utc
            self._validade = not_after.strftime("%d/%m/%Y")
        except AttributeError:
            try:
                self._validade = certificate.not_valid_after.strftime("%d/%m/%Y")
            except Exception:
                self._validade = "?"

        # CNPJ (extraído do CN ou SAN)
        self._cnpj = self._extrair_cnpj(self._titular)
        if not self._cnpj:
            self._cnpj = self._extrair_cnpj_san(certificate)

    def _carregar_pyopenssl(self, pfx_data: bytes):
        p12 = openssl_crypto.load_pkcs12(pfx_data, self._senha_bytes)

        self._cert_pem = openssl_crypto.dump_certificate(
            openssl_crypto.FILETYPE_PEM, p12.get_certificate()
        )
        self._key_pem = openssl_crypto.dump_privatekey(
            openssl_crypto.FILETYPE_PEM, p12.get_privatekey()
        )

        cert = p12.get_certificate()
        subject = cert.get_subject()
        self._titular = subject.CN or "Desconhecido"
        self._validade = datetime.strptime(
            cert.get_notAfter().decode(), "%Y%m%d%H%M%SZ"
        ).strftime("%d/%m/%Y")
        self._cnpj = self._extrair_cnpj(self._titular)

    @staticmethod
    def _extrair_cnpj(texto: str) -> str:
        """Extrai CNPJ (14 dígitos) de uma string como o CN do certificado."""
        if not texto:
            return ""
        nums = re.findall(r"\d{14}", texto.replace(".", "").replace("/", "").replace("-", ""))
        return nums[0] if nums else ""

    @staticmethod
    def _extrair_cnpj_san(certificate) -> str:
        """Tenta extrair CNPJ das SANs ou OtherName do certificado."""
        try:
            san = certificate.extensions.get_extension_for_class(
                x509.SubjectAlternativeName
            )
            for name in san.value:
                txt = str(name.value) if hasattr(name, "value") else ""
                cnpj = CertificadoDigital._extrair_cnpj(txt)
                if cnpj:
                    return cnpj
        except Exception:
            pass
        return ""

    def info(self) -> dict:
        return {
            "titular": self._titular,
            "cnpj": self._cnpj,
            "validade": self._validade,
            "arquivo": os.path.basename(self.caminho),
        }

    def pem_cert(self) -> bytes:
        return self._cert_pem

    def pem_key(self) -> bytes:
        return self._key_pem

    def exportar_pem_temp(self):
        """
        Cria arquivos PEM temporários e retorna (cert_path, key_path).
        O chamador é responsável por remover os arquivos após uso.
        """
        cert_file = tempfile.NamedTemporaryFile(delete=False, suffix=".cert.pem")
        cert_file.write(self._cert_pem)
        cert_file.close()

        key_file = tempfile.NamedTemporaryFile(delete=False, suffix=".key.pem")
        key_file.write(self._key_pem)
        key_file.close()

        return cert_file.name, key_file.name

    @property
    def titular(self) -> str:
        return self._titular

    @property
    def cnpj(self) -> str:
        return self._cnpj

    @property
    def validade(self) -> str:
        return self._validade
