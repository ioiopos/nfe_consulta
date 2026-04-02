"""
Cliente SEFAZ - Web Service NFeDistribuicaoDFe
Consulta NF-e Modelo 55 destinadas via Ambiente Nacional (AN).

Referência: NT 2014.002 v1.10
URL Produção:  https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx
URL Homolog.:  https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx

NOTAS SOBRE O FORMATO DE RESPOSTA DO SEFAZ:
  O SEFAZ retorna o retDistDFeInt de duas formas possíveis:
  a) Como elemento XML filho direto no Body SOAP
  b) Como texto dentro de <nfeDistDFeInteresseResult> (forma mais comum na prática)
  Este cliente trata ambas as formas com cascata de estratégias.
"""

import os
import re
import gzip
import base64
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import requests
import urllib3

from src.certificado import CertificadoDigital
from src.parser import ParserNFe

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL_PRODUCAO    = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"
URL_HOMOLOGACAO = "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

NS_NFE  = "http://www.portalfiscal.inf.br/nfe"
SOAPACTION = "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDistDFeInteresse"

DEBUG_DIR = Path(__file__).parent.parent / "logs"
DEBUG_DIR.mkdir(exist_ok=True)


class SefazClient:
    def __init__(self, certificado, ambiente: int = 1, cuf_autor: str = "43"):
        self.cert      = certificado
        self.ambiente  = ambiente
        self.cuf_autor = cuf_autor
        self.url       = URL_PRODUCAO if ambiente == 1 else URL_HOMOLOGACAO
        # Detecta se o certificado é do Windows Store (não precisa de PEM)
        self._modo_winhttp = getattr(certificado, "_winhttp_cn", None) is not None

    def consultar_distribuicao(self, cnpj: str, ultimo_nsu: int = 0) -> dict:
        xml_body      = self._montar_xml(cnpj, ultimo_nsu)
        soap_envelope = self._montar_soap(xml_body)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")

        if self._modo_winhttp:
            return self._consultar_winhttp(soap_envelope, ts, ultimo_nsu)
        else:
            return self._consultar_requests(soap_envelope, ts, ultimo_nsu)

    def _consultar_winhttp(self, soap_envelope: str, ts: str, ultimo_nsu: int) -> dict:
        """Usa WinHTTP COM — para certificados não-exportáveis do Windows Store."""
        try:
            import pythoncom
            from src.win_cert_store import requisicao_winhttp

            # CoInitialize obrigatório em threads que usam COM
            pythoncom.CoInitialize()
            try:
                cn = self.cert._winhttp_cn
                response_text = requisicao_winhttp(
                    self.url,
                    soap_envelope.encode("utf-8"),
                    cn,
                    headers={"SOAPAction": SOAPACTION}
                )
                (DEBUG_DIR / f"resposta_{ts}.xml").write_text(response_text, encoding="utf-8")
                return self._processar_resposta(response_text, ultimo_nsu)
            finally:
                pythoncom.CoUninitialize()

        except Exception as e:
            return self._erro(str(e), "WINHTTP_ERRO", ultimo_nsu)

    def _consultar_requests(self, soap_envelope: str, ts: str, ultimo_nsu: int) -> dict:
        """Usa requests + PEM — para certificados carregados de arquivo .pfx."""
        cert_path = key_path = None
        try:
            cert_path, key_path = self.cert.exportar_pem_temp()

            response = requests.post(
                self.url,
                data=soap_envelope.encode("utf-8"),
                headers={
                    "Content-Type": "application/soap+xml; charset=utf-8",
                    "SOAPAction": SOAPACTION,
                },
                cert=(cert_path, key_path),
                verify=False,
                timeout=60,
            )

            (DEBUG_DIR / f"resposta_{ts}.xml").write_text(response.text, encoding="utf-8")

            if response.status_code != 200:
                return self._erro(
                    f"HTTP {response.status_code}: {response.text[:300]}",
                    str(response.status_code), ultimo_nsu, response.text
                )

            return self._processar_resposta(response.text, ultimo_nsu)

        finally:
            for p in [cert_path, key_path]:
                if p and os.path.exists(p):
                    try:
                        os.unlink(p)
                    except Exception:
                        pass

    def _montar_xml(self, cnpj: str, ultimo_nsu: int) -> str:
        # cUFAutor: código IBGE de 2 dígitos da UF do autor (tipo TCodUfIBGE)
        # NÃO usar 91 aqui — 91 é código do Ambiente Nacional, inválido neste campo
        return (
            f'<distDFeInt xmlns="http://www.portalfiscal.inf.br/nfe" versao="1.01">'
            f'<tpAmb>{self.ambiente}</tpAmb>'
            f'<cUFAutor>{self.cuf_autor}</cUFAutor>'
            f'<CNPJ>{cnpj}</CNPJ>'
            f'<distNSU><ultNSU>{ultimo_nsu:015d}</ultNSU></distNSU>'
            f'</distDFeInt>'
        )

    def _montar_soap(self, xml_body: str) -> str:
        # CRÍTICO: nfeDadosMsg deve ter o namespace explícito declarado
        # O XML interno vai como conteúdo de texto da tag (já é string, não aninhado)
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

    def _processar_resposta(self, xml_texto: str, ultimo_nsu: int) -> dict:
        ret_xml = self._extrair_ret_dist(xml_texto)
        if ret_xml is None:
            return self._erro(
                "retDistDFeInt não localizado. Verifique logs/resposta_*.xml",
                "SEM_RETORNO", ultimo_nsu, xml_texto
            )

        try:
            ret = ET.fromstring(ret_xml)
        except ET.ParseError as e:
            return self._erro(f"Erro ao parsear retDistDFeInt: {e}", "XML_INVALIDO", ultimo_nsu, xml_texto)

        def tag(nome):
            el = ret.find(f"{{{NS_NFE}}}{nome}")
            if el is None:
                el = ret.find(nome)
            return el.text.strip() if el is not None and el.text else ""

        cstat = tag("cStat")
        xmot  = tag("xMotivo")

        if cstat == "656":
            return {"status": "consumo_indevido", "codigo": cstat, "mensagem": xmot,
                    "notas": [], "max_nsu": ultimo_nsu, "xml_bruto": xml_texto}

        max_nsu   = self._int_tag(ret, "maxNSU", ultimo_nsu)
        ult_nsu   = self._int_tag(ret, "ultNSU", ultimo_nsu)
        nsu_final = max(max_nsu, ult_nsu, ultimo_nsu)

        # Localiza lote — sem ou com namespace
        lote = ret.find(f"{{{NS_NFE}}}loteDistDFeInt") or ret.find("loteDistDFeInt")

        docs = list(lote.iter()) if lote is not None else []
        doc_zips = [el for el in docs if el.tag.endswith("docZip")]

        if not doc_zips:
            return {"status": "sem_novos", "codigo": cstat, "mensagem": xmot,
                    "notas": [], "max_nsu": nsu_final, "xml_bruto": xml_texto}

        parser = ParserNFe()
        notas  = []
        for doc_el in doc_zips:
            nsu_doc = 0
            schema  = ""
            try:
                nsu_doc = int(doc_el.get("NSU") or doc_el.get("nsu") or "0")
                schema  = doc_el.get("schema") or doc_el.get("Schema") or ""
                raw     = (doc_el.text or "").strip()
                if not raw:
                    continue
                xml_doc = gzip.decompress(base64.b64decode(raw)).decode("utf-8")
                nota = parser.parse_documento(xml_doc, schema, nsu_doc)
                nota["xml_raw"] = xml_doc
                notas.append(nota)
                nsu_final = max(nsu_final, nsu_doc)
            except Exception as e:
                notas.append({"nsu": nsu_doc, "chave": "", "emitente": "", "valor": "",
                               "emissao": "", "situacao": "Erro leitura", "schema": schema,
                               "erro_parse": str(e), "xml_raw": ""})

        return {"status": "ok", "codigo": cstat, "mensagem": xmot,
                "notas": notas, "max_nsu": nsu_final, "xml_bruto": xml_texto}

    def _extrair_ret_dist(self, xml_texto: str):
        """
        Cascata de 3 estratégias para extrair o retDistDFeInt.
        """
        # 1) Parse SOAP direto — itera todos os elementos
        try:
            root = ET.fromstring(xml_texto)
            for el in root.iter():
                if el.tag.endswith("retDistDFeInt"):
                    return ET.tostring(el, encoding="unicode")
        except ET.ParseError:
            pass

        # 2) Conteúdo de texto em *Result (XML escapado dentro da tag)
        try:
            root = ET.fromstring(xml_texto)
            for el in root.iter():
                if not ("Result" in el.tag or "result" in el.tag):
                    continue
                inner = (el.text or "").strip()
                if not inner.startswith("<"):
                    continue
                try:
                    inner_root = ET.fromstring(inner)
                    for iel in inner_root.iter():
                        if iel.tag.endswith("retDistDFeInt"):
                            return ET.tostring(iel, encoding="unicode")
                    if inner_root.tag.endswith("retDistDFeInt"):
                        return inner
                except ET.ParseError:
                    pass
        except ET.ParseError:
            pass

        # 3) Regex — último recurso
        m = re.search(r'(<retDistDFeInt[\s\S]*?</retDistDFeInt>)', xml_texto, re.DOTALL)
        if m:
            frag = m.group(1)
            if "xmlns" not in frag:
                frag = frag.replace("<retDistDFeInt", f'<retDistDFeInt xmlns="{NS_NFE}"', 1)
            return frag

        return None

    def _int_tag(self, el: ET.Element, tag: str, default: int) -> int:
        found = el.find(f"{{{NS_NFE}}}{tag}") or el.find(tag)
        if found is not None and found.text:
            try:
                return int(found.text.strip())
            except ValueError:
                pass
        return default

    @staticmethod
    def _erro(mensagem, codigo, ultimo_nsu, xml_bruto=""):
        return {"status": "erro", "codigo": codigo, "mensagem": mensagem,
                "notas": [], "max_nsu": ultimo_nsu, "xml_bruto": xml_bruto}
