"""
Cliente SEFAZ - NFeRecepcaoEvento (Manifestação do Destinatário)
e download do XML completo via NFeDistribuicaoDFe (consNSU).

Tipos de evento de manifestação:
  210200 - Confirmação da Operação      (conclusivo - mercadoria recebida)
  210210 - Ciência da Operação          (declara ciência, não conclusivo)
  210220 - Desconhecimento da Operação  (conclusivo - não reconhece)
  210240 - Operação não Realizada       (conclusivo - requer justificativa)

URL Produção:    https://www.nfe.fazenda.gov.br/RecepcaoEvento/RecepcaoEvento.asmx
URL Homologação: https://hom.nfe.fazenda.gov.br/RecepcaoEvento/RecepcaoEvento.asmx
"""

import os
import re
import gzip
import base64
import random
import tempfile
import xml.etree.ElementTree as ET
from datetime import datetime, timezone, timedelta

import requests
import urllib3

from src.certificado import CertificadoDigital

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ─────────────────────────────────────────────────────────────────────────────
URL_EVENTO_PROD = "https://www.nfe.fazenda.gov.br/RecepcaoEvento/RecepcaoEvento.asmx"
URL_EVENTO_HOM  = "https://hom.nfe.fazenda.gov.br/RecepcaoEvento/RecepcaoEvento.asmx"

URL_DIST_PROD   = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"
URL_DIST_HOM    = "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

NS_NFE = "http://www.portalfiscal.inf.br/nfe"

TIPOS_EVENTO = {
    "210200": "Confirmacao da Operacao",
    "210210": "Ciencia da Operacao",
    "210220": "Desconhecimento da Operacao",
    "210240": "Operacao nao Realizada",
}

# cOrgao = 91 para o Ambiente Nacional (sempre para manifestação)
C_ORGAO_AN = "91"


class EventoClient:
    """Envia eventos de Manifestação do Destinatário e baixa XMLs completos."""

    def __init__(self, certificado, ambiente: int = 1):
        self.cert     = certificado
        self.ambiente = ambiente
        self.url_ev   = URL_EVENTO_PROD if ambiente == 1 else URL_EVENTO_HOM
        self.url_dist = URL_DIST_PROD   if ambiente == 1 else URL_DIST_HOM
        self._modo_winhttp = getattr(certificado, "_winhttp_cn", None) is not None

    # ── Manifestação ─────────────────────────────────────────────────────────

    def manifestar(self, cnpj: str, chave: str, tp_evento: str,
                   justificativa: str = "", n_seq: int = 1) -> dict:
        xml_evento   = self._montar_xml_evento(cnpj, chave, tp_evento, justificativa, n_seq)
        soap_action  = ("http://www.portalfiscal.inf.br/nfe/wsdl/"
                        "NFeRecepcaoEvento4/nfeRecepcaoEvento")

        if self._modo_winhttp:
            return self._post_winhttp(
                self.url_ev, xml_evento, self.cert._winhttp_cn,
                soap_action, self._processar_retorno_evento, assinar=True
            )

        cert_path = key_path = None
        try:
            cert_path, key_path = self.cert.exportar_pem_temp()
            xml_assinado = self._assinar(xml_evento, cert_path, key_path)
            soap = self._montar_soap_evento(xml_assinado)
            resp = requests.post(
                self.url_ev,
                data=soap.encode("utf-8"),
                headers={
                    "Content-Type": "application/soap+xml; charset=utf-8",
                    "SOAPAction": soap_action,
                },
                cert=(cert_path, key_path),
                verify=False, timeout=60,
            )
            if resp.status_code != 200:
                return {"status": "erro", "codigo": str(resp.status_code),
                        "mensagem": f"HTTP {resp.status_code}", "protocolo": ""}
            return self._processar_retorno_evento(resp.text)
        finally:
            for p in [cert_path, key_path]:
                if p and os.path.exists(p):
                    try: os.unlink(p)
                    except Exception: pass

    # ── Download XML ─────────────────────────────────────────────────────────

    def baixar_xml(self, cnpj: str, nsu: int) -> dict:
        """Baixa o XML completo de uma NF-e via consNSU (NSU específico)."""
        xml_req     = self._montar_cons_nsu(cnpj, nsu)
        soap        = self._montar_soap_dist(xml_req)
        soap_action = ("http://www.portalfiscal.inf.br/nfe/wsdl/"
                       "NFeDistribuicaoDFe/nfeDistDFeInteresse")

        if self._modo_winhttp:
            return self._post_winhttp(
                self.url_dist, soap, self.cert._winhttp_cn,
                soap_action, lambda r: self._processar_retorno_xml(r, nsu)
            )

        cert_path = key_path = None
        try:
            cert_path, key_path = self.cert.exportar_pem_temp()
            resp = requests.post(
                self.url_dist, data=soap.encode("utf-8"),
                headers={"Content-Type": "application/soap+xml; charset=utf-8",
                         "SOAPAction": soap_action},
                cert=(cert_path, key_path), verify=False, timeout=60,
            )
            return self._processar_retorno_xml(resp.text, nsu)
        finally:
            for p in [cert_path, key_path]:
                if p and os.path.exists(p):
                    try: os.unlink(p)
                    except Exception: pass

    # ── WinHTTP COM (certificados não-exportáveis) ────────────────────────────

    def _post_winhttp(self, url, xml_or_soap, cn, soap_action, processar_fn, assinar=False):
        """POST via WinHTTP COM com CoInitialize — funciona em qualquer thread."""
        try:
            import pythoncom
            from src.win_cert_store import requisicao_winhttp

            pythoncom.CoInitialize()
            try:
                body = xml_or_soap if isinstance(xml_or_soap, bytes) else xml_or_soap.encode("utf-8")
                response_text = requisicao_winhttp(url, body, cn,
                                                   headers={"SOAPAction": soap_action})
                return processar_fn(response_text)
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            return {"status": "erro", "codigo": "WINHTTP_ERRO",
                    "mensagem": str(e), "protocolo": "", "xml_str": ""}

    # ── Montagem de XMLs ─────────────────────────────────────────────────────

    def _montar_xml_evento(self, cnpj, chave, tp_evento,
                            justificativa, n_seq) -> str:
        dh_evento = self._dh_agora()
        desc = TIPOS_EVENTO.get(tp_evento, "Ciencia da Operacao")
        id_evento = f"ID{tp_evento}{chave}{n_seq:02d}"
        id_lote   = str(random.randint(1, 999999999999999)).zfill(15)

        det_extra = ""
        if tp_evento == "210240" and justificativa:
            det_extra = f"\n    <xJust>{justificativa[:255]}</xJust>"

        return f"""<?xml version="1.0" encoding="UTF-8"?>
<envEvento xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.00">
  <idLote>{id_lote}</idLote>
  <evento xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.00">
    <infEvento Id="{id_evento}">
      <cOrgao>{C_ORGAO_AN}</cOrgao>
      <tpAmb>{self.ambiente}</tpAmb>
      <CNPJ>{cnpj}</CNPJ>
      <chNFe>{chave}</chNFe>
      <dhEvento>{dh_evento}</dhEvento>
      <tpEvento>{tp_evento}</tpEvento>
      <nSeqEvento>{n_seq}</nSeqEvento>
      <verEvento>1.00</verEvento>
      <detEvento versao="1.00">
        <descEvento>{desc}</descEvento>{det_extra}
      </detEvento>
    </infEvento>
  </evento>
</envEvento>"""

    def _montar_cons_nsu(self, cnpj: str, nsu: int) -> str:
        nsu_fmt = f"{nsu:015d}"
        return f"""<distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">
  <tpAmb>{self.ambiente}</tpAmb>
  <cUFAutor>91</cUFAutor>
  <CNPJ>{cnpj}</CNPJ>
  <consNSU>
    <NSU>{nsu_fmt}</NSU>
  </consNSU>
</distDFeInt>"""

    def _montar_soap_evento(self, xml_body: str) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<soap12:Envelope'
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
            ' xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
            ' xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
            '<soap12:Body>'
            '<nfeRecepcaoEvento xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeRecepcaoEvento4">'
            '<nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeRecepcaoEvento4">'
            f'{xml_body}'
            '</nfeDadosMsg>'
            '</nfeRecepcaoEvento>'
            '</soap12:Body>'
            '</soap12:Envelope>'
        )

    def _montar_soap_dist(self, xml_body: str) -> str:
        return (
            '<?xml version="1.0" encoding="UTF-8"?>'
            '<soap12:Envelope'
            ' xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"'
            ' xmlns:xsd="http://www.w3.org/2001/XMLSchema"'
            ' xmlns:soap12="http://www.w3.org/2003/05/soap-envelope">'
            '<soap12:Body>'
            '<nfeDistDFeInteresse xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">'
            '<nfeDadosMsg xmlns="http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe">'
            f'{xml_body}'
            '</nfeDadosMsg>'
            '</nfeDistDFeInteresse>'
            '</soap12:Body>'
            '</soap12:Envelope>'
        )

    # ── Assinatura ────────────────────────────────────────────────────────────

    def _assinar(self, xml_str: str, cert_path: str, key_path: str) -> str:
        """Assina o XML usando signxml (lxml). Fallback para str sem assinatura."""
        try:
            from src.assinatura import AssinadorNFe
            cert_pem = open(cert_path, "rb").read()
            key_pem  = open(key_path,  "rb").read()
            assinador = AssinadorNFe(cert_pem, key_pem)
            return assinador.assinar_evento(xml_str)
        except ImportError:
            # signxml/lxml não instalado — retorna sem assinatura
            # O SEFAZ vai rejeitar, mas permite testar o fluxo
            return xml_str

    # ── Processamento de retornos ────────────────────────────────────────────

    def _processar_retorno_evento(self, xml_resp: str) -> dict:
        try:
            root = ET.fromstring(xml_resp)
        except ET.ParseError:
            return {"status": "erro", "codigo": "?", "mensagem": "XML inválido", "protocolo": ""}

        def t(tag):
            el = root.find(f".//{{{NS_NFE}}}{tag}")
            return el.text.strip() if el is not None and el.text else ""

        cstat    = t("cStat")
        xmot     = t("xMotivo")
        nprot    = t("nProt")
        dh_reg   = t("dhRegEvento")

        if cstat in ("135", "136"):
            return {
                "status": "ok",
                "codigo": cstat,
                "mensagem": xmot,
                "protocolo": nprot,
                "dh_registro": dh_reg,
            }
        else:
            return {
                "status": "erro",
                "codigo": cstat,
                "mensagem": xmot,
                "protocolo": "",
            }

    def _processar_retorno_xml(self, xml_resp: str, nsu: int) -> dict:
        try:
            root = ET.fromstring(xml_resp)
        except ET.ParseError:
            return {"status": "erro", "mensagem": "Resposta XML inválida", "xml_str": ""}

        doc_zip = root.find(f".//{{{NS_NFE}}}docZip")
        if doc_zip is None or not doc_zip.text:
            cstat_el = root.find(f".//{{{NS_NFE}}}cStat")
            xmot_el  = root.find(f".//{{{NS_NFE}}}xMotivo")
            cstat = cstat_el.text.strip() if cstat_el is not None else "?"
            xmot  = xmot_el.text.strip()  if xmot_el  is not None else "?"
            return {
                "status": "sem_xml",
                "codigo": cstat,
                "mensagem": xmot,
                "xml_str": "",
            }

        try:
            zipped  = base64.b64decode(doc_zip.text)
            xml_str = gzip.decompress(zipped).decode("utf-8")
            return {"status": "ok", "xml_str": xml_str, "nsu": nsu}
        except Exception as e:
            return {"status": "erro", "mensagem": str(e), "xml_str": ""}

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _dh_agora() -> str:
        """Data/hora atual no formato exigido pelo SEFAZ: AAAA-MM-DDThh:mm:ss-03:00"""
        tz_br = timezone(timedelta(hours=-3))
        now   = datetime.now(tz_br)
        return now.strftime("%Y-%m-%dT%H:%M:%S-03:00")
