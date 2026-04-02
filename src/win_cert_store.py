"""
Windows Certificate Store - Leitura e exportação de certificados instalados.

Usa ssl.enum_certificates() para listar e crypt32.dll via ctypes para exportar
o PFX (certificado + chave privada) em memória — sem arquivos temporários.

Funciona APENAS no Windows. No Linux/macOS retorna lista vazia sem erro.
"""

import ssl
import sys
import re
import ctypes
import ctypes.wintypes as wintypes
from datetime import datetime, timezone

from cryptography import x509
from cryptography.hazmat.primitives.serialization import pkcs12, Encoding, PrivateFormat, NoEncryption

IS_WINDOWS = sys.platform == "win32"

# ─── Constantes crypt32 ───────────────────────────────────────────────────────
CRYPT_EXPORTABLE                 = 0x00000001
CERT_STORE_ADD_USE_EXISTING      = 2
CERT_CLOSE_STORE_FORCE_FLAG      = 1
X509_ASN_ENCODING                = 0x00000001
PKCS_7_ASN_ENCODING              = 0x00010000
CERT_STORE_PROV_MEMORY           = b"Memory"
CERT_STORE_PROV_SYSTEM_A         = 9
CERT_SYSTEM_STORE_CURRENT_USER   = 0x00010000
CERT_FIND_ANY                    = 0x00000000


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


def listar_certificados_windows() -> list:
    """
    Lista certificados válidos com chave privada no store MY do usuário atual.
    Retorna lista de dicts: {titular, cnpj, validade, validade_dt, der}
    """
    if not IS_WINDOWS:
        return []

    resultado = []
    try:
        for der_raw, encoding, trust in ssl.enum_certificates("MY"):
            if encoding != "x509_asn":
                continue
            try:
                der = bytes(der_raw)
                cert = x509.load_der_x509_certificate(der)
                cn   = _extrair_cn(cert)
                cnpj = _extrair_cnpj(cn)

                # Validade
                try:
                    not_after = cert.not_valid_after_utc
                    now = datetime.now(timezone.utc)
                except AttributeError:
                    not_after = cert.not_valid_after
                    now = datetime.now()

                if not_after < now:
                    continue

                resultado.append({
                    "titular":     cn,
                    "cnpj":        cnpj,
                    "validade":    not_after.strftime("%d/%m/%Y"),
                    "validade_dt": not_after,
                    "der":         der,
                })
            except Exception:
                continue
    except Exception:
        pass

    return resultado


def exportar_pem_do_store(der_cert: bytes):
    """
    Dado o DER de um certificado instalado no Windows MY store,
    exporta como PFX em memória e retorna (cert_pem_bytes, key_pem_bytes).
    """
    if not IS_WINDOWS:
        raise RuntimeError("Disponível apenas no Windows.")

    pfx_data = _exportar_pfx_por_der(der_cert)
    if not pfx_data:
        raise RuntimeError(
            "Falha ao exportar o certificado.\n\n"
            "Possíveis causas:\n"
            "• O certificado foi instalado como não-exportável\n"
            "• Permissão negada pelo Windows\n\n"
            "Use a opção de carregar o arquivo .pfx manualmente."
        )

    # PFX exportado com senha vazia
    try:
        private_key, certificate, _ = pkcs12.load_key_and_certificates(pfx_data, b"")
    except Exception:
        # Tenta com None como senha
        private_key, certificate, _ = pkcs12.load_key_and_certificates(pfx_data, None)

    cert_pem = certificate.public_bytes(Encoding.PEM)
    key_pem  = private_key.private_bytes(
        encoding=Encoding.PEM,
        format=PrivateFormat.PKCS8,
        encryption_algorithm=NoEncryption(),
    )
    return cert_pem, key_pem


def _exportar_pfx_por_der(der_cert: bytes) -> bytes:
    """
    Implementação correta da exportação PFX via crypt32.dll:
    1. Abre o store MY do usuário
    2. Encontra o certificado pelo conteúdo DER iterando o store
    3. Cria um store temporário em memória
    4. Adiciona apenas este certificado ao store temporário
    5. Exporta o store temporário como PFX (inclui chave privada automaticamente)
    """
    crypt32 = ctypes.windll.crypt32

    # 1. Abre o store MY
    store_my = crypt32.CertOpenStore(
        CERT_STORE_PROV_SYSTEM_A,
        0,
        None,
        CERT_SYSTEM_STORE_CURRENT_USER,
        b"MY"
    )
    if not store_my:
        raise RuntimeError(f"Não foi possível abrir o store MY. Erro: {ctypes.GetLastError()}")

    cert_ctx = None
    try:
        # 2. Itera o store e encontra o certificado pelo DER
        cert_ctx = _encontrar_cert_por_der(crypt32, store_my, der_cert)
        if not cert_ctx:
            raise RuntimeError("Certificado não encontrado no store MY do Windows.")

        # 3. Cria store temporário em memória
        store_mem = crypt32.CertOpenStore(
            CERT_STORE_PROV_MEMORY,
            0, None, 0, None
        )
        if not store_mem:
            raise RuntimeError("Não foi possível criar store temporário em memória.")

        try:
            # 4. Adiciona o certificado (e sua chave privada vinculada) ao store temporário
            ok = crypt32.CertAddCertificateContextToStore(
                store_mem,
                cert_ctx,
                CERT_STORE_ADD_USE_EXISTING,
                None
            )
            if not ok:
                raise RuntimeError(f"Falha ao adicionar certificado ao store temp. Erro: {ctypes.GetLastError()}")

            # 5. Exporta o store temporário como PFX
            # Primeira chamada: descobre o tamanho do buffer necessário
            pfx_blob = DATA_BLOB()
            pfx_blob.cbData = 0
            pfx_blob.pbData = None

            ok = crypt32.PFXExportCertStoreEx(
                store_mem,
                ctypes.byref(pfx_blob),
                ctypes.c_wchar_p(""),   # senha vazia
                None,
                CRYPT_EXPORTABLE
            )

            if not ok or pfx_blob.cbData == 0:
                err = ctypes.GetLastError()
                raise RuntimeError(f"PFXExportCertStoreEx (tamanho) falhou. Erro Windows: {err}")

            # Segunda chamada: aloca buffer e exporta de fato
            buf = (ctypes.c_ubyte * pfx_blob.cbData)()
            pfx_blob.pbData = buf

            ok2 = crypt32.PFXExportCertStoreEx(
                store_mem,
                ctypes.byref(pfx_blob),
                ctypes.c_wchar_p(""),
                None,
                CRYPT_EXPORTABLE
            )

            if not ok2:
                err = ctypes.GetLastError()
                raise RuntimeError(f"PFXExportCertStoreEx (dados) falhou. Erro Windows: {err}")

            return bytes(buf[:pfx_blob.cbData])

        finally:
            crypt32.CertCloseStore(store_mem, CERT_CLOSE_STORE_FORCE_FLAG)
    finally:
        if cert_ctx:
            crypt32.CertFreeCertificateContext(cert_ctx)
        crypt32.CertCloseStore(store_my, 0)


class _CERT_CONTEXT(ctypes.Structure):
    _fields_ = [
        ("dwCertEncodingType", wintypes.DWORD),
        ("pbCertEncoded",      ctypes.POINTER(ctypes.c_ubyte)),
        ("cbCertEncoded",      wintypes.DWORD),
        ("pCertInfo",          ctypes.c_void_p),
        ("hCertStore",         ctypes.c_void_p),
    ]


def _encontrar_cert_por_der(crypt32, store, der_alvo: bytes):
    """Itera o store e retorna o contexto do certificado que bate com o DER."""
    ctx_prev = None
    while True:
        ctx = crypt32.CertEnumCertificatesInStore(store, ctx_prev)
        if not ctx:
            break
        ctx_prev = ctx
        ctx_ptr = ctypes.cast(ctx, ctypes.POINTER(_CERT_CONTEXT))
        size = ctx_ptr.contents.cbCertEncoded
        der_atual = bytes(ctx_ptr.contents.pbCertEncoded[:size])
        if der_atual == der_alvo:
            # Duplica o contexto para que o iterador não o libere
            dup = crypt32.CertDuplicateCertificateContext(ctx)
            return dup
    return None


# ── Helpers ────────────────────────────────────────────────────────────────────

def _extrair_cn(cert) -> str:
    try:
        attrs = cert.subject.get_attributes_for_oid(x509.NameOID.COMMON_NAME)
        return attrs[0].value if attrs else ""
    except Exception:
        return ""


def _extrair_cnpj(texto: str) -> str:
    if not texto:
        return ""
    limpo = re.sub(r'[\.\-\/\s]', '', texto)
    nums  = re.findall(r'\d{14}', limpo)
    return nums[0] if nums else ""
