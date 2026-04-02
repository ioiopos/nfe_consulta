"""
Windows Certificate Store — listagem e requisição HTTP via WinHTTP COM.

Para certificados NÃO-EXPORTÁVEIS (recomendação de segurança):
  Usa WinHttp.WinHttpRequest.5.1 via win32com.client, que delega a autenticação
  mútua TLS diretamente ao SChannel do Windows — a chave privada nunca sai do store.

Para certificados EXPORTÁVEIS:
  Tenta exportar via crypt32.dll como fallback.

Funciona APENAS no Windows. Em outros SOs retorna lista vazia sem erro.
"""

import ssl
import sys
import re
import ctypes
import ctypes.wintypes as wintypes
from datetime import datetime, timezone

from cryptography import x509
from cryptography.hazmat.primitives.serialization import (
    pkcs12, Encoding, PrivateFormat, NoEncryption
)

IS_WINDOWS = sys.platform == "win32"

CRYPT_EXPORTABLE               = 0x00000001
CERT_STORE_ADD_USE_EXISTING    = 2
CERT_CLOSE_STORE_FORCE_FLAG    = 1
CERT_STORE_PROV_MEMORY         = b"Memory"
CERT_STORE_PROV_SYSTEM_A       = 9
CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000


class DATA_BLOB(ctypes.Structure):
    _fields_ = [
        ("cbData", wintypes.DWORD),
        ("pbData", ctypes.POINTER(ctypes.c_ubyte)),
    ]


class _CERT_CONTEXT(ctypes.Structure):
    _fields_ = [
        ("dwCertEncodingType", wintypes.DWORD),
        ("pbCertEncoded",      ctypes.POINTER(ctypes.c_ubyte)),
        ("cbCertEncoded",      wintypes.DWORD),
        ("pCertInfo",          ctypes.c_void_p),
        ("hCertStore",         ctypes.c_void_p),
    ]


# ─── API pública ──────────────────────────────────────────────────────────────

def listar_certificados_windows() -> list:
    """
    Lista certificados válidos (não expirados) no store MY do usuário atual.
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
                der  = bytes(der_raw)
                cert = x509.load_der_x509_certificate(der)
                cn   = _extrair_cn(cert)
                cnpj = _extrair_cnpj(cn)

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


def requisicao_winhttp(url: str, soap_body: bytes, cn_certificado: str,
                       headers: dict = None) -> str:
    """
    Faz POST via WinHttp.WinHttpRequest.5.1 COM.
    Certificado identificado pelo CN — chave nunca é exportada.
    """
    if not IS_WINDOWS:
        raise RuntimeError("WinHTTP disponível apenas no Windows.")

    try:
        import win32com.client
    except ImportError:
        raise RuntimeError(
            "pywin32 não instalado.\n"
            "Execute: pip install pywin32\n\n"
            "O pywin32 é necessário para usar certificados não-exportáveis do Windows."
        )

    try:
        req = win32com.client.Dispatch("WinHttp.WinHttpRequest.5.1")

        req.Open("POST", url, False)  # False = síncrono

        # Ignora erros de CA da Fazenda (não está no store de raízes padrão)
        # WinHttpRequestOption_SslErrorIgnoreFlags = 4
        # Valor 0x3300 ignora: CA inválida + certificado revogado
        # Sintaxe correta no win32com: req.Option[indice] = valor  OU  SetProperty
        try:
            req.Option[4] = 0x3300
        except Exception:
            # Fallback: alguns ambientes exigem a sintaxe de setter direta
            try:
                req.SetProperty("Option", 4, 0x3300)
            except Exception:
                pass  # Continua sem ignorar erros de CA

        # Seleciona o certificado pelo CN no store Pessoal do usuário atual
        cert_string = f"CURRENT_USER\\MY\\{cn_certificado}"
        req.SetClientCertificate(cert_string)

        # Headers
        req.SetRequestHeader("Content-Type", "application/soap+xml; charset=utf-8")
        if headers:
            for k, v in headers.items():
                req.SetRequestHeader(k, v)

        # Envia
        req.Send(soap_body)

        status = req.Status
        if status != 200:
            raise RuntimeError(f"HTTP {status}: {req.StatusText}")

        return req.ResponseText

    except RuntimeError:
        raise
    except Exception as e:
        raise RuntimeError(f"WinHTTP falhou: {e}") from e


def exportar_pem_do_store(der_cert: bytes):
    """
    Tenta exportar certificado + chave privada como PEM via crypt32.
    Funciona APENAS se o certificado foi instalado como exportável.
    Para certificados não-exportáveis, use requisicao_winhttp().
    """
    if not IS_WINDOWS:
        raise RuntimeError("Disponível apenas no Windows.")

    pfx_data = _exportar_pfx(der_cert)

    try:
        private_key, certificate, _ = pkcs12.load_key_and_certificates(pfx_data, b"")
    except Exception:
        private_key, certificate, _ = pkcs12.load_key_and_certificates(pfx_data, None)

    cert_pem = certificate.public_bytes(Encoding.PEM)
    key_pem  = private_key.private_bytes(
        encoding=Encoding.PEM,
        format=PrivateFormat.PKCS8,
        encryption_algorithm=NoEncryption(),
    )
    return cert_pem, key_pem


# ─── Exportação PFX via crypt32 (apenas para exportáveis) ────────────────────

def _init_crypt32():
    c = ctypes.windll.crypt32

    c.CertOpenStore.restype  = ctypes.c_void_p
    c.CertOpenStore.argtypes = [
        ctypes.c_char_p, wintypes.DWORD, ctypes.c_void_p,
        wintypes.DWORD, ctypes.c_void_p,
    ]
    c.CertEnumCertificatesInStore.restype  = ctypes.c_void_p
    c.CertEnumCertificatesInStore.argtypes = [ctypes.c_void_p, ctypes.c_void_p]
    c.CertDuplicateCertificateContext.restype  = ctypes.c_void_p
    c.CertDuplicateCertificateContext.argtypes = [ctypes.c_void_p]
    c.CertFreeCertificateContext.restype  = wintypes.BOOL
    c.CertFreeCertificateContext.argtypes = [ctypes.c_void_p]
    c.CertAddCertificateContextToStore.restype  = wintypes.BOOL
    c.CertAddCertificateContextToStore.argtypes = [
        ctypes.c_void_p, ctypes.c_void_p,
        wintypes.DWORD, ctypes.POINTER(ctypes.c_void_p),
    ]
    c.PFXExportCertStoreEx.restype  = wintypes.BOOL
    c.PFXExportCertStoreEx.argtypes = [
        ctypes.c_void_p, ctypes.POINTER(DATA_BLOB),
        wintypes.LPCWSTR, ctypes.c_void_p, wintypes.DWORD,
    ]
    c.CertCloseStore.restype  = wintypes.BOOL
    c.CertCloseStore.argtypes = [ctypes.c_void_p, wintypes.DWORD]
    return c


def _exportar_pfx(der_cert: bytes) -> bytes:
    c = _init_crypt32()

    store_my = c.CertOpenStore(
        CERT_STORE_PROV_SYSTEM_A, 0, None,
        CERT_SYSTEM_STORE_CURRENT_USER, b"MY"
    )
    if not store_my:
        raise RuntimeError(f"Não foi possível abrir o store MY. Erro: {ctypes.GetLastError()}")

    cert_ctx = store_mem = None
    try:
        cert_ctx = _encontrar_cert(c, store_my, der_cert)
        if not cert_ctx:
            raise RuntimeError("Certificado não encontrado no store MY.")

        store_mem = c.CertOpenStore(CERT_STORE_PROV_MEMORY, 0, None, 0, None)
        if not store_mem:
            raise RuntimeError("Falha ao criar store temporário em memória.")

        ok = c.CertAddCertificateContextToStore(
            store_mem, cert_ctx, CERT_STORE_ADD_USE_EXISTING, None
        )
        if not ok:
            raise RuntimeError(
                f"Falha ao copiar certificado (erro {ctypes.GetLastError()}).\n"
                "O certificado provavelmente foi instalado como NÃO-EXPORTÁVEL.\n"
                "Use a opção de arquivo .pfx."
            )

        # Descobre tamanho
        blob = DATA_BLOB()
        blob.cbData = 0
        blob.pbData = None
        ok = c.PFXExportCertStoreEx(store_mem, ctypes.byref(blob), "", None, CRYPT_EXPORTABLE)
        if not ok or blob.cbData == 0:
            raise RuntimeError(
                f"PFXExportCertStoreEx falhou (erro {ctypes.GetLastError()}).\n"
                "Certificado não-exportável — use o arquivo .pfx."
            )

        # Exporta
        buf = (ctypes.c_ubyte * blob.cbData)()
        blob.pbData = buf
        ok = c.PFXExportCertStoreEx(store_mem, ctypes.byref(blob), "", None, CRYPT_EXPORTABLE)
        if not ok:
            raise RuntimeError(f"PFXExportCertStoreEx (dados) falhou. Erro: {ctypes.GetLastError()}")

        return bytes(buf[:blob.cbData])
    finally:
        if cert_ctx:
            c.CertFreeCertificateContext(cert_ctx)
        if store_mem:
            c.CertCloseStore(store_mem, CERT_CLOSE_STORE_FORCE_FLAG)
        c.CertCloseStore(store_my, 0)


def _encontrar_cert(c, store, der_alvo: bytes):
    ctx_prev = None
    while True:
        ctx = c.CertEnumCertificatesInStore(store, ctx_prev)
        if not ctx:
            break
        ctx_prev = ctx
        ptr  = ctypes.cast(ctx, ctypes.POINTER(_CERT_CONTEXT))
        size = ptr.contents.cbCertEncoded
        der  = bytes(ptr.contents.pbCertEncoded[:size])
        if der == der_alvo:
            return c.CertDuplicateCertificateContext(ctx)
    return None


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
