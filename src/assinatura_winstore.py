"""
Assinatura XML de eventos NF-e usando Windows Certificate Store (CAPI).

Para certificados NAO-EXPORTAVEIS — usa CryptSignHash via ctypes.

CORRECAO CRITICA:
  CryptSignHashW com dwFlags=0 (nao CRYPT_NOHASHOID) inclui o DigestInfo
  no padding PKCS#1 v1.5 — obrigatorio para RSA-SHA1 valido no XMLDSig.
  Usar CRYPT_NOHASHOID=1 omite o DigestInfo e gera assinatura invalida.
"""

import base64
import hashlib
import ctypes
import ctypes.wintypes as wintypes
from lxml import etree

# Constantes CAPI
X509_ASN_ENCODING              = 0x00000001
PKCS_7_ASN_ENCODING            = 0x00010000
CERT_STORE_PROV_SYSTEM         = ctypes.c_void_p(10)
CERT_SYSTEM_STORE_CURRENT_USER = 0x00010000
CERT_FIND_SUBJECT_STR_W        = 0x00080004
CERT_FIND_EXISTING             = 0x000D0000
CALG_SHA1                      = 0x00008004
HP_HASHVAL                     = 0x0002
# IMPORTANTE: 0 (nao CRYPT_NOHASHOID) para incluir DigestInfo no PKCS#1
SIGN_FLAGS                     = 0

NS_DS  = "http://www.w3.org/2000/09/xmldsig#"
NS_NFE = "http://www.portalfiscal.inf.br/nfe"

_c = ctypes.windll.crypt32
_a = ctypes.windll.advapi32

# Tipos
_c.CertOpenStore.restype  = ctypes.c_void_p
_c.CertOpenStore.argtypes = [ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p,
                              wintypes.DWORD, ctypes.c_wchar_p]
_c.CertFindCertificateInStore.restype  = ctypes.c_void_p
_c.CertFindCertificateInStore.argtypes = [ctypes.c_void_p, wintypes.DWORD, wintypes.DWORD,
                                          wintypes.DWORD, ctypes.c_void_p, ctypes.c_void_p]
_c.CertCreateCertificateContext.restype  = ctypes.c_void_p
_c.CertCreateCertificateContext.argtypes = [wintypes.DWORD, ctypes.c_char_p, wintypes.DWORD]
_c.CertFreeCertificateContext.restype  = wintypes.BOOL
_c.CertFreeCertificateContext.argtypes = [ctypes.c_void_p]
_c.CertCloseStore.restype  = wintypes.BOOL
_c.CertCloseStore.argtypes = [ctypes.c_void_p, wintypes.DWORD]
_c.CryptAcquireCertificatePrivateKey.restype  = wintypes.BOOL
_c.CryptAcquireCertificatePrivateKey.argtypes = [
    ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_void_p),
    ctypes.POINTER(wintypes.DWORD),
    ctypes.POINTER(wintypes.BOOL)]
_a.CryptCreateHash.restype  = wintypes.BOOL
_a.CryptCreateHash.argtypes = [ctypes.c_void_p, wintypes.DWORD, ctypes.c_void_p,
                                wintypes.DWORD, ctypes.POINTER(ctypes.c_void_p)]
_a.CryptSetHashParam.restype  = wintypes.BOOL
_a.CryptSetHashParam.argtypes = [ctypes.c_void_p, wintypes.DWORD,
                                  ctypes.c_char_p, wintypes.DWORD]
_a.CryptSignHashW.restype  = wintypes.BOOL
_a.CryptSignHashW.argtypes = [ctypes.c_void_p, wintypes.DWORD, ctypes.c_wchar_p,
                               wintypes.DWORD, ctypes.c_char_p,
                               ctypes.POINTER(wintypes.DWORD)]
_a.CryptDestroyHash.restype  = wintypes.BOOL
_a.CryptDestroyHash.argtypes = [ctypes.c_void_p]
_a.CryptReleaseContext.restype  = wintypes.BOOL
_a.CryptReleaseContext.argtypes = [ctypes.c_void_p, wintypes.DWORD]


class CERT_CONTEXT(ctypes.Structure):
    _fields_ = [("dwCertEncodingType", wintypes.DWORD),
                 ("pbCertEncoded",      ctypes.POINTER(ctypes.c_ubyte)),
                 ("cbCertEncoded",      wintypes.DWORD),
                 ("pCertInfo",          ctypes.c_void_p),
                 ("hCertStore",         ctypes.c_void_p)]


def _abrir_store():
    h = _c.CertOpenStore(CERT_STORE_PROV_SYSTEM,
                         X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
                         None, CERT_SYSTEM_STORE_CURRENT_USER, "MY")
    if not h:
        raise RuntimeError(f"CertOpenStore falhou: {ctypes.GetLastError()}")
    return h


def _encontrar_cert(store, cn: str, der: bytes):
    """Busca o certificado no store: DER exato -> nome -> parte do nome."""
    # 1) DER exato
    if der:
        try:
            tmp = _c.CertCreateCertificateContext(
                X509_ASN_ENCODING | PKCS_7_ASN_ENCODING, der, len(der))
            if tmp:
                h = _c.CertFindCertificateInStore(
                    store, X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
                    0, CERT_FIND_EXISTING, tmp, None)
                _c.CertFreeCertificateContext(tmp)
                if h:
                    return h
        except Exception:
            pass

    # 2) CN completo (ex: "PEDRO HENRIQUE...:13667015000150")
    h = _c.CertFindCertificateInStore(
        store, X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
        0, CERT_FIND_SUBJECT_STR_W, ctypes.c_wchar_p(cn), None)
    if h:
        return h

    # 3) So o nome (sem CNPJ)
    nome = cn.split(":")[0].strip() if ":" in cn else cn
    h = _c.CertFindCertificateInStore(
        store, X509_ASN_ENCODING | PKCS_7_ASN_ENCODING,
        0, CERT_FIND_SUBJECT_STR_W, ctypes.c_wchar_p(nome), None)
    if h:
        return h

    raise RuntimeError(
        f"Certificado nao encontrado no store MY.\n"
        f"Buscado: '{cn}'\n"
        f"Verifique: Gerenciar certificados de usuario > Pessoal > Certificados")


def _cert_der_bytes(cert_h) -> bytes:
    ctx = ctypes.cast(cert_h, ctypes.POINTER(CERT_CONTEXT)).contents
    return bytes(ctx.pbCertEncoded[:ctx.cbCertEncoded])


def _rsa_sha1_sign(cert_h, data_hash: bytes) -> bytes:
    """
    Assina usando RSA-SHA1 PKCS#1 v1.5 via CAPI.
    dwFlags=0 garante inclusao do DigestInfo — obrigatorio para XMLDSig.
    """
    h_prov    = ctypes.c_void_p()
    dw_key    = wintypes.DWORD()
    free_prov = wintypes.BOOL()

    if not _c.CryptAcquireCertificatePrivateKey(
            cert_h, 0, None,
            ctypes.byref(h_prov), ctypes.byref(dw_key), ctypes.byref(free_prov)):
        raise RuntimeError(f"CryptAcquireCertificatePrivateKey: {ctypes.GetLastError()}")
    try:
        h_hash = ctypes.c_void_p()
        if not _a.CryptCreateHash(h_prov, CALG_SHA1, None, 0, ctypes.byref(h_hash)):
            raise RuntimeError(f"CryptCreateHash: {ctypes.GetLastError()}")
        try:
            if not _a.CryptSetHashParam(h_hash, HP_HASHVAL, data_hash, 0):
                raise RuntimeError(f"CryptSetHashParam: {ctypes.GetLastError()}")

            sig_len = wintypes.DWORD(0)
            # Primeira chamada: descobre o tamanho (sem CRYPT_NOHASHOID)
            _a.CryptSignHashW(h_hash, dw_key.value, None, SIGN_FLAGS,
                              None, ctypes.byref(sig_len))

            buf = ctypes.create_string_buffer(sig_len.value)
            if not _a.CryptSignHashW(h_hash, dw_key.value, None, SIGN_FLAGS,
                                     buf, ctypes.byref(sig_len)):
                raise RuntimeError(f"CryptSignHashW: {ctypes.GetLastError()}")

            # CAPI retorna little-endian; RSA PKCS#1 precisa big-endian
            return bytes(buf.raw[:sig_len.value])[::-1]
        finally:
            _a.CryptDestroyHash(h_hash)
    finally:
        if free_prov.value:
            _a.CryptReleaseContext(h_prov, 0)


def _c14n(el) -> bytes:
    return etree.tostring(el, method="c14n", exclusive=True, with_comments=False)


def assinar_evento_winstore(xml_str: str, cn: str, der: bytes = b"") -> str:
    """
    Assina o XML de evento NF-e com o certificado do Windows Store.
    Usa RSA-SHA1 PKCS#1 v1.5 via CAPI com DigestInfo correto.
    Retorna o XML com a assinatura XMLDSig inserida no elemento <evento>.
    """
    store = cert_h = None
    try:
        store  = _abrir_store()
        cert_h = _encontrar_cert(store, cn, der)
        cert_b64 = base64.b64encode(_cert_der_bytes(cert_h)).decode()

        root = etree.fromstring(xml_str.encode("utf-8"))

        # Localiza infEvento
        inf = next((el for el in root.iter() if el.tag.endswith("infEvento")), None)
        if inf is None:
            raise ValueError("infEvento nao encontrado")
        ref_id = inf.get("Id", "")

        # DigestValue do infEvento canonizado
        digest_b64 = base64.b64encode(hashlib.sha1(_c14n(inf)).digest()).decode()

        # SignedInfo
        si_str = (
            f'<SignedInfo xmlns="{NS_DS}">'
            '<CanonicalizationMethod Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>'
            '<SignatureMethod Algorithm="http://www.w3.org/2000/09/xmldsig#rsa-sha1"/>'
            f'<Reference URI="#{ref_id}">'
            '<Transforms>'
            '<Transform Algorithm="http://www.w3.org/2000/09/xmldsig#enveloped-signature"/>'
            '<Transform Algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315"/>'
            '</Transforms>'
            '<DigestMethod Algorithm="http://www.w3.org/2000/09/xmldsig#sha1"/>'
            f'<DigestValue>{digest_b64}</DigestValue>'
            '</Reference>'
            '</SignedInfo>'
        )
        si_el   = etree.fromstring(si_str.encode("utf-8"))
        si_hash = hashlib.sha1(_c14n(si_el)).digest()

        # Assina o hash do SignedInfo
        sig_bytes = _rsa_sha1_sign(cert_h, si_hash)
        sig_b64   = base64.b64encode(sig_bytes).decode()

        # Monta Signature completa
        sig_el = etree.fromstring((
            f'<Signature xmlns="{NS_DS}">'
            f'{si_str}'
            f'<SignatureValue>{sig_b64}</SignatureValue>'
            f'<KeyInfo><X509Data>'
            f'<X509Certificate>{cert_b64}</X509Certificate>'
            f'</X509Data></KeyInfo>'
            f'</Signature>'
        ).encode("utf-8"))

        # Insere no elemento <evento>
        ev_el = next((el for el in root.iter()
                      if el.tag.endswith("}evento") or el.tag == "evento"), None)
        if ev_el is None:
            raise ValueError("elemento <evento> nao encontrado")
        ev_el.append(sig_el)

        # SEM declaração XML — será inserido dentro do SOAP envelope
        return etree.tostring(root, encoding="unicode")

    finally:
        if cert_h: _c.CertFreeCertificateContext(cert_h)
        if store:  _c.CertCloseStore(store, 0)
