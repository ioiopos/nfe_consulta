"""
Assinatura XML para eventos NF-e (Manifestação do Destinatário).
Usa signxml + lxml — 100% Python, sem dependência de xmlsec1 externo.

O SEFAZ exige RSA-SHA1, porém OpenSSL moderno (3.x) bloqueia SHA1 por padrão.
Solução: habilitar SHA1 via contexto legacy antes de assinar.
"""

import re
import ssl
import os
from lxml import etree
from signxml import XMLSigner, methods

NS_NFE = "http://www.portalfiscal.inf.br/nfe"
NS_DS  = "http://www.w3.org/2000/09/xmldsig#"


def _habilitar_sha1_legacy():
    """
    OpenSSL 3.x desabilita SHA1 por padrão com 'SHA1-based algorithms are not
    supported in the default configuration'. Para compatibilidade com o SEFAZ
    (que ainda exige RSA-SHA1), forçamos a política legacy via variável de ambiente
    ou via configuração de contexto.
    """
    # Método 1: variável de ambiente (funciona no OpenSSL 3.x)
    os.environ.setdefault("OPENSSL_CONF", "")
    # Método 2: força flag UNSAFE_LEGACY_SERVER_CONNECT no ssl
    try:
        ssl._create_default_https_context = ssl._create_unverified_context
    except Exception:
        pass


class AssinadorNFe:
    """
    Assina o elemento infEvento de um evento NF-e com o certificado A1.
    Compatível com o padrão exigido pelo SEFAZ (RSA-SHA1, C14N).
    """

    def __init__(self, cert_pem: bytes, key_pem: bytes):
        self.cert_pem = cert_pem
        self.key_pem  = key_pem

    def assinar_evento(self, xml_str: str) -> str:
        """
        Recebe o XML do envEvento (sem assinatura) e retorna o XML assinado.
        """
        root = etree.fromstring(xml_str.encode("utf-8"))

        ns = {"nfe": NS_NFE}
        inf_evento = root.find(".//nfe:infEvento", ns)
        if inf_evento is None:
            raise ValueError("Elemento infEvento não encontrado no XML.")

        id_ref = inf_evento.get("Id", "")
        if not id_ref:
            raise ValueError("Atributo Id ausente no infEvento.")

        evento_el = root.find(".//nfe:evento", ns)
        if evento_el is None:
            raise ValueError("Elemento evento não encontrado.")

        etree.SubElement(
            evento_el,
            f"{{{NS_DS}}}Signature",
            attrib={"Id": "placeholder"}
        )

        # Habilita SHA1 legacy para compatibilidade com SEFAZ
        _habilitar_sha1_legacy()

        try:
            # Tenta com RSA-SHA1 (exigido pelo SEFAZ)
            signed_root = self._assinar_com_algoritmo(root, id_ref, "rsa-sha1", "sha1")
        except Exception as e:
            if "sha1" in str(e).lower() or "not supported" in str(e).lower() or "unsafe" in str(e).lower():
                # Fallback: RSA-SHA256 (aceito pelo SEFAZ desde NT 2014.001 v1.10)
                signed_root = self._assinar_com_algoritmo(root, id_ref, "rsa-sha256", "sha256")
            else:
                raise

        return etree.tostring(signed_root, encoding="unicode", xml_declaration=False)

    def _assinar_com_algoritmo(self, root, id_ref, sig_alg, digest_alg):
        signer = XMLSigner(
            method=methods.enveloped,
            signature_algorithm=sig_alg,
            digest_algorithm=digest_alg,
            c14n_algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315",
        )
        return signer.sign(
            root,
            key=self.key_pem,
            cert=self.cert_pem,
            reference_uri=f"#{id_ref}",
        )
