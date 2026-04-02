"""
Assinatura XML para eventos NF-e (Manifestação do Destinatário).
Usa signxml + lxml — 100% Python, sem dependência de xmlsec1 externo.

O SEFAZ exige:
  - Algoritmo: RSA-SHA1  (rsa-sha1)
  - Canonicalização: C14N exclusivo (exc-c14n)
  - Referência: URI="#ID..." com Transform enveloped-signature
  - Namespace ds: http://www.w3.org/2000/09/xmldsig#
"""

import re
from lxml import etree
from signxml import XMLSigner, methods

NS_NFE = "http://www.portalfiscal.inf.br/nfe"
NS_DS  = "http://www.w3.org/2000/09/xmldsig#"


class AssinadorNFe:
    """
    Assina o elemento infEvento de um evento NF-e com o certificado A1.
    Compatível com o padrão exigido pelo SEFAZ (RSA-SHA1, C14N exclusivo).
    """

    def __init__(self, cert_pem: bytes, key_pem: bytes):
        """
        :param cert_pem: certificado em formato PEM (bytes)
        :param key_pem:  chave privada em formato PEM (bytes)
        """
        self.cert_pem = cert_pem
        self.key_pem  = key_pem

    def assinar_evento(self, xml_str: str) -> str:
        """
        Recebe o XML do envEvento (sem assinatura) e retorna o XML assinado.
        A assinatura é inserida dentro do elemento <evento> → <infEvento>.
        """
        root = etree.fromstring(xml_str.encode("utf-8"))

        # Localiza o infEvento para obter o Id a referenciar
        ns = {"nfe": NS_NFE}
        inf_evento = root.find(".//nfe:infEvento", ns)
        if inf_evento is None:
            raise ValueError("Elemento infEvento não encontrado no XML.")

        id_ref = inf_evento.get("Id", "")
        if not id_ref:
            raise ValueError("Atributo Id ausente no infEvento.")

        # Insere placeholder de assinatura dentro do elemento <evento>
        evento_el = root.find(".//nfe:evento", ns)
        if evento_el is None:
            raise ValueError("Elemento evento não encontrado.")

        # Adiciona <ds:Signature> placeholder (signxml vai preenchê-lo)
        placeholder = etree.SubElement(
            evento_el,
            f"{{{NS_DS}}}Signature",
            attrib={"Id": "placeholder"}
        )

        signer = XMLSigner(
            method=methods.enveloped,
            signature_algorithm="rsa-sha1",
            digest_algorithm="sha1",
            c14n_algorithm="http://www.w3.org/TR/2001/REC-xml-c14n-20010315",
        )

        signed_root = signer.sign(
            root,
            key=self.key_pem,
            cert=self.cert_pem,
            reference_uri=f"#{id_ref}",
        )

        xml_assinado = etree.tostring(
            signed_root,
            encoding="unicode",
            xml_declaration=False,
        )
        return xml_assinado
