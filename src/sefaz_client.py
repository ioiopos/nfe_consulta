"""
Cliente SEFAZ - NFeDistribuicaoDFe
NT 2014.002 v1.30 (Fevereiro 2026) — implementação fiel à legislação.

REGRAS CRÍTICAS DA NT (seções 3.4, 3.5, 3.11.4):

1. NSU é sequencial — SEMPRE usar o ultNSU retornado na consulta anterior.
   Nunca pular, nunca recuar, nunca usar maxNSU como próximo ponto.

2. Novo usuário (CNPJ nunca consultou distNSU):
   - Primeiro acesso retorna cStat=137. É NORMAL. O SEFAZ começa a gerar NSU a partir daí.
   - Aguardar 1 hora. Próxima consulta começa a trazer documentos.
   - NÃO fazer consultas extras de sondagem (consomem o token de 60 min).

3. cStat=137 com ultNSU==maxNSU: chegou ao teto. Aguardar 1h.
   cStat=137 com maxNSU>ultNSU: há mais blocos adiante. Continuar com ultNSU.
   cStat=138: documentos encontrados. Próxima consulta com ultNSU retornado.

4. Qualquer consulta dentro de 1h após cStat=137 = cStat=656 (bloqueio por 1h).
   Se receber 656 antes de completar 1h, o TIMER É ZERADO e recomeça.

5. Persistir NSU: sempre salvar o ultNSU retornado (não maxNSU).
   Salvar mesmo quando cStat=137 — indica até onde o SEFAZ processou.
"""

import os
import re
import gzip
import base64
import html
import xml.etree.ElementTree as ET
from datetime import datetime
from pathlib import Path

import requests
import urllib3

from src.parser import ParserNFe

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

URL_PRODUCAO    = "https://www1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"
URL_HOMOLOGACAO = "https://hom1.nfe.fazenda.gov.br/NFeDistribuicaoDFe/NFeDistribuicaoDFe.asmx"

NS_NFE     = "http://www.portalfiscal.inf.br/nfe"
SOAPACTION = "http://www.portalfiscal.inf.br/nfe/wsdl/NFeDistribuicaoDFe/nfeDistDFeInteresse"

TIMEOUT_SEFAZ     = 90    # SEFAZ pode demorar 30-45s em carga
DELAY_ENTRE_LOTES = 1.0   # evita 656 por sobrecarga

_BASE_DIR = Path(os.environ.get("ABBAS_NFE_DIR", Path(__file__).parent.parent))
DEBUG_DIR = _BASE_DIR / "logs"
DEBUG_DIR.mkdir(exist_ok=True)

# Tabela oficial de cStat (NT 2014.002 seção 4)
CSTAT_DESCRICAO = {
    "137": ("sem_novos",       "Nenhum documento localizado para este NSU"),
    "138": ("ok",              "Documento(s) localizado(s)"),
    "656": ("consumo_indevido","Consumo indevido — aguarde 60 minutos"),
    "108": ("aviso",           "Serviço paralisado momentaneamente"),
    "109": ("aviso",           "Serviço paralisado sem previsão"),
    "214": ("erro",            "Tamanho da mensagem excedeu o limite (10 KB)"),
    "215": ("erro",            "Falha no schema XML — verifique o XML enviado"),
    "217": ("erro",            "NF-e inexistente para a chave de acesso informada"),
    "252": ("erro",            "Ambiente informado diverge do ambiente de recebimento"),
    "280": ("erro",            "Certificado transmissor inválido"),
    "281": ("erro",            "Certificado transmissor vencido"),
    "283": ("erro",            "Certificado — erro na cadeia de certificação"),
    "284": ("erro",            "Certificado transmissor revogado"),
    "285": ("erro",            "Certificado transmissor difere ICP-Brasil"),
    "286": ("erro",            "Certificado — erro no acesso à LCR"),
    "402": ("erro",            "XML com codificação diferente de UTF-8"),
    "404": ("erro",            "Uso de prefixo de namespace não permitido"),
    "472": ("erro",            "CPF consultado difere do CPF do certificado digital"),
    "473": ("erro",            "Certificado sem CNPJ ou CPF"),
    "489": ("erro",            "CNPJ informado inválido (DV ou zeros)"),
    "490": ("erro",            "CPF informado inválido (DV ou zeros)"),
    "589": ("erro",            "NSU informado é superior ao maior NSU do Ambiente Nacional"),
    "593": ("erro",            "CNPJ-base consultado difere do CNPJ-base do certificado"),
    "632": ("erro",            "NF-e fora do prazo — não disponível para download (> 90 dias)"),
    "640": ("erro",            "CNPJ/CPF não possui permissão para consultar esta NF-e"),
    "656": ("consumo_indevido","Consumo indevido — aguarde 60 minutos"),
    "999": ("erro",            "Erro não catalogado"),
}


def descrever_cstat(cstat: str, xmotivo: str) -> tuple:
    """Retorna (status, descricao_humana) para um cStat."""
    if cstat in CSTAT_DESCRICAO:
        st, descr = CSTAT_DESCRICAO[cstat]
        return st, f"[{cstat}] {descr}"
    # Não mapeado — usa xMotivo do SEFAZ
    if cstat.startswith("1"):
        return "ok",   f"[{cstat}] {xmotivo}"
    if cstat.startswith(("2", "3", "4", "5", "6")):
        return "erro", f"[{cstat}] {xmotivo}"
    return "aviso", f"[{cstat}] {xmotivo}"


def _tag_flex(el: ET.Element, nome: str) -> str:
    """Busca tag com ou sem namespace, ou por iteração."""
    for ns in [f"{{{NS_NFE}}}", ""]:
        found = el.find(f"{ns}{nome}")
        if found is not None and found.text:
            return found.text.strip()
    for child in el.iter():
        local = child.tag.split("}")[-1] if "}" in child.tag else child.tag
        if local == nome and child.text:
            return child.text.strip()
    return ""


def _int_flex(el: ET.Element, nome: str, default: int) -> int:
    val = _tag_flex(el, nome)
    try:
        return int(val) if val else default
    except ValueError:
        return default


class SefazClient:
    def __init__(self, certificado, ambiente: int = 1, cuf_autor: str = "43"):
        self.cert          = certificado
        self.ambiente      = ambiente
        self.cuf_autor     = cuf_autor
        self.url           = URL_PRODUCAO if ambiente == 1 else URL_HOMOLOGACAO
        self._modo_winhttp = getattr(certificado, "_winhttp_cn", None) is not None
        self._diag_log     = []

    def consultar_distribuicao(self, cnpj: str, ultimo_nsu: int = 0) -> dict:
        """
        Consulta distNSU conforme NT 2014.002 seção 3.5.
        Envia ultNSU e recebe lote de até 50 documentos com NSU superior.

        ATENÇÃO: cada chamada consome 1 token por CNPJ.
        Após cStat=137: aguardar 60 minutos antes de nova chamada.
        """
        xml_body      = self._montar_xml_dist(cnpj, ultimo_nsu)
        soap_envelope = self._montar_soap(xml_body)
        ts            = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        self._diag_log = []
        self._diag(f"distNSU: CNPJ={cnpj} ultNSU={ultimo_nsu:015d}")

        if self._modo_winhttp:
            return self._consultar_winhttp(soap_envelope, ts, ultimo_nsu)
        else:
            return self._consultar_requests(soap_envelope, ts, ultimo_nsu)

    # ── Transporte ─────────────────────────────────────────────────────────

    def _consultar_winhttp(self, soap_envelope: str, ts: str, ultimo_nsu: int) -> dict:
        try:
            import pythoncom
            from src.win_cert_store import requisicao_winhttp
            pythoncom.CoInitialize()
            try:
                import time
                t0 = time.time()
                resp = requisicao_winhttp(
                    self.url, soap_envelope.encode("utf-8"),
                    self.cert._winhttp_cn, headers={"SOAPAction": SOAPACTION})
                elapsed = time.time() - t0
                self._diag(f"WinHTTP: {elapsed:.1f}s — {len(resp)} bytes")

                if not resp or len(resp) < 100:
                    self._salvar_diag(ts, "Resposta vazia WinHTTP", resp)
                    return self._erro(
                        "WinHTTP retornou resposta vazia.\n"
                        "Verifique o certificado e a conexão.",
                        "WINHTTP_VAZIO", ultimo_nsu)

                self._salvar_resp(ts, resp)
                return self._processar_resposta(resp, ultimo_nsu, ts)
            finally:
                pythoncom.CoUninitialize()
        except Exception as e:
            msg = str(e)
            self._diag(f"EXCEÇÃO WinHTTP: {msg}")
            dica = ""
            if "12002" in msg or "timeout" in msg.lower():
                dica = "\nO SEFAZ não respondeu no tempo. Tente novamente."
            elif "12029" in msg or "connect" in msg.lower():
                dica = "\nSem conexão com o SEFAZ. Verifique a internet."
            elif "12175" in msg or "certificate" in msg.lower():
                dica = "\nErro de certificado. Verifique validade e instalação."
            self._salvar_diag(ts, msg, "")
            return self._erro(f"WinHTTP falhou: {msg}{dica}", "WINHTTP_ERRO", ultimo_nsu)

    def _consultar_requests(self, soap_envelope: str, ts: str, ultimo_nsu: int) -> dict:
        cert_path = key_path = None
        try:
            cert_path, key_path = self.cert.exportar_pem_temp()
            import time
            t0 = time.time()
            response = requests.post(
                self.url,
                data=soap_envelope.encode("utf-8"),
                headers={"Content-Type": "application/soap+xml; charset=utf-8",
                         "SOAPAction": SOAPACTION},
                cert=(cert_path, key_path),
                verify=False,
                timeout=TIMEOUT_SEFAZ,
            )
            elapsed = time.time() - t0
            self._diag(f"HTTP {response.status_code} em {elapsed:.1f}s — {len(response.text)} bytes")
            self._salvar_resp(ts, response.text)

            if response.status_code != 200:
                return self._erro(
                    f"HTTP {response.status_code}: {response.text[:300]}",
                    str(response.status_code), ultimo_nsu, response.text)
            return self._processar_resposta(response.text, ultimo_nsu, ts)

        except requests.exceptions.Timeout:
            return self._erro(
                f"Timeout após {TIMEOUT_SEFAZ}s — SEFAZ sem resposta. Tente novamente.",
                "TIMEOUT", ultimo_nsu)
        except requests.exceptions.ConnectionError as e:
            return self._erro(
                f"Sem conexão com o SEFAZ: {e}",
                "CONN_ERROR", ultimo_nsu)
        except Exception as e:
            return self._erro(str(e), "REQUESTS_ERRO", ultimo_nsu)
        finally:
            for p in [cert_path, key_path]:
                if p and os.path.exists(p):
                    try: os.unlink(p)
                    except: pass

    # ── Montagem XML ────────────────────────────────────────────────────────

    def _montar_xml_dist(self, cnpj: str, ultimo_nsu: int) -> str:
        """distNSU — NT seção 3.4.1a. cUFAutor = código IBGE da UF (não 91)."""
        return (
            f'<distDFeInt xmlns="{NS_NFE}" versao="1.01">'
            f'<tpAmb>{self.ambiente}</tpAmb>'
            f'<cUFAutor>{self.cuf_autor}</cUFAutor>'
            f'<CNPJ>{cnpj}</CNPJ>'
            f'<distNSU><ultNSU>{ultimo_nsu:015d}</ultNSU></distNSU>'
            f'</distDFeInt>'
        )

    def _montar_soap(self, xml_body: str) -> str:
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

    # ── Processamento da resposta ───────────────────────────────────────────

    def _processar_resposta(self, xml_texto: str, ultimo_nsu: int, ts: str = "") -> dict:
        ret_xml = self._extrair_ret_dist(xml_texto)
        if ret_xml is None:
            self._salvar_diag(ts, "retDistDFeInt não encontrado", xml_texto)
            return self._erro(
                "retDistDFeInt não localizado na resposta SEFAZ.\n"
                f"Diagnóstico em logs/diag_{ts}.txt",
                "SEM_RETORNO", ultimo_nsu, xml_texto)

        try:
            ret = ET.fromstring(ret_xml)
        except ET.ParseError as e:
            return self._erro(f"Erro ao parsear retDistDFeInt: {e}",
                              "XML_INVALIDO", ultimo_nsu, xml_texto)

        cstat = _tag_flex(ret, "cStat")
        xmot  = _tag_flex(ret, "xMotivo")

        # Se cStat vazio, tenta sem namespace
        if not cstat:
            ret_sem_ns = re.sub(r'\s*xmlns[^"]*"[^"]*"', '', ret_xml)
            ret_sem_ns = re.sub(r"\s*xmlns[^']*'[^']*'", '', ret_sem_ns)
            try:
                ret2  = ET.fromstring(ret_sem_ns)
                cstat = _tag_flex(ret2, "cStat")
                xmot  = _tag_flex(ret2, "xMotivo")
                if cstat:
                    ret = ret2
            except ET.ParseError:
                pass

        if not cstat:
            self._salvar_diag(ts, f"cStat vazio\n{ret_xml[:500]}", xml_texto)
            return self._erro(
                "Resposta SEFAZ sem cStat — formato inesperado.\n"
                f"Diagnóstico em logs/diag_{ts}.txt",
                "CSTAT_VAZIO", ultimo_nsu, xml_texto)

        status_cstat, descr_humana = descrever_cstat(cstat, xmot)
        self._diag(f"cStat={cstat} | {descr_humana}")

        # NT B08/B09: ultNSU = último pesquisado; maxNSU = teto do AN para este CNPJ
        max_nsu = _int_flex(ret, "maxNSU", ultimo_nsu)
        ult_nsu = _int_flex(ret, "ultNSU", ultimo_nsu)
        self._diag(f"ultNSU={ult_nsu} maxNSU={max_nsu}")

        # NT 3.11.4: 656 = consumo indevido
        if cstat == "656":
            return {"status": "consumo_indevido", "codigo": cstat,
                    "mensagem": descr_humana,
                    "notas": [], "ult_nsu": ult_nsu, "max_nsu": max_nsu,
                    "tem_mais": False}

        if status_cstat == "erro":
            return {"status": "erro", "codigo": cstat,
                    "mensagem": descr_humana,
                    "notas": [], "ult_nsu": ult_nsu, "max_nsu": max_nsu,
                    "tem_mais": False}

        # Localiza lote de documentos
        lote = None
        for el in ret.iter():
            if el.tag.endswith("loteDistDFeInt"):
                lote = el
                break

        doc_zips = []
        if lote is not None:
            for el in lote.iter():
                if el.tag.endswith("docZip"):
                    doc_zips.append(el)

        self._diag(f"docZip: {len(doc_zips)}")

        # NT 3.5: cStat=137 → sem docs neste bloco de NSU
        # tem_mais = maxNSU > ultNSU (há mais blocos adiante)
        # NT 3.11.4.2: próxima consulta usa ultNSU (não maxNSU)
        if not doc_zips:
            tem_mais = (max_nsu > ult_nsu)
            return {"status": "sem_novos", "codigo": cstat,
                    "mensagem": descr_humana,
                    "notas": [], "ult_nsu": ult_nsu, "max_nsu": max_nsu,
                    "tem_mais": tem_mais}

        parser = ParserNFe()
        notas  = []
        nsu_ultimo_doc = ult_nsu
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
                nota    = parser.parse_documento(xml_doc, schema, nsu_doc)
                nota["xml_raw"] = xml_doc
                notas.append(nota)
                nsu_ultimo_doc = max(nsu_ultimo_doc, nsu_doc)
            except Exception as e:
                notas.append({"nsu": nsu_doc, "chave": "", "emitente": "",
                               "valor": "", "emissao": "", "situacao": "Erro leitura",
                               "schema": schema, "erro_parse": str(e), "xml_raw": ""})

        # cStat=138: tem mais documentos (continua loop com ultNSU)
        tem_mais = (cstat == "138")

        return {"status": "ok", "codigo": cstat, "mensagem": descr_humana,
                "tem_mais": tem_mais,
                "notas": notas,
                "ult_nsu": ult_nsu,            # usar este na próxima consulta
                "max_nsu": max_nsu,            # teto do AN
                "nsu_ultimo_doc": nsu_ultimo_doc}

    # ── Extração retDistDFeInt (4 estratégias em cascata) ──────────────────

    def _extrair_ret_dist(self, xml_texto: str):
        # 1) Parse SOAP direto
        try:
            root = ET.fromstring(xml_texto)
            for el in root.iter():
                if el.tag.endswith("retDistDFeInt"):
                    return ET.tostring(el, encoding="unicode")
        except ET.ParseError:
            pass

        # 2) Conteúdo escapado dentro de tag *Result
        try:
            root = ET.fromstring(xml_texto)
            for el in root.iter():
                if not any(k in el.tag for k in ["Result", "result", "Response"]):
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

        # 3) Regex direto
        m = re.search(r'(<retDistDFeInt[\s\S]*?</retDistDFeInt>)', xml_texto, re.DOTALL)
        if m:
            frag = m.group(1)
            if "xmlns" not in frag:
                frag = frag.replace("<retDistDFeInt",
                                    f'<retDistDFeInt xmlns="{NS_NFE}"', 1)
            return frag

        # 4) html.unescape (WinHTTP às vezes entrega entidades HTML)
        if "&lt;retDistDFeInt" in xml_texto:
            decoded = html.unescape(xml_texto)
            m2 = re.search(r'(<retDistDFeInt[\s\S]*?</retDistDFeInt>)', decoded, re.DOTALL)
            if m2:
                frag = m2.group(1)
                if "xmlns" not in frag:
                    frag = frag.replace("<retDistDFeInt",
                                        f'<retDistDFeInt xmlns="{NS_NFE}"', 1)
                self._diag("Estratégia 4 (html.unescape) usada")
                return frag

        return None

    # ── Diagnóstico ────────────────────────────────────────────────────────

    def _diag(self, msg: str):
        self._diag_log.append(f"[{datetime.now().strftime('%H:%M:%S.%f')[:-3]}] {msg}")

    def _salvar_resp(self, ts: str, texto: str):
        try:
            (DEBUG_DIR / f"resp_{ts}.xml").write_text(texto, encoding="utf-8")
        except Exception:
            pass

    def _salvar_diag(self, ts: str, erro: str, resp: str):
        try:
            linhas = ["=== DIAGNÓSTICO SEFAZ ===",
                      f"Timestamp: {ts}", f"Erro: {erro}", "",
                      "=== LOG INTERNO ==="] + self._diag_log + [
                      "", "=== RESPOSTA BRUTA (primeiros 2000 chars) ===",
                      (resp or "")[:2000]]
            (DEBUG_DIR / f"diag_{ts}.txt").write_text("\n".join(linhas), encoding="utf-8")
        except Exception:
            pass

    @staticmethod
    def _erro(mensagem, codigo, ultimo_nsu, xml_bruto=""):
        return {"status": "erro", "codigo": codigo, "mensagem": mensagem,
                "notas": [], "ult_nsu": ultimo_nsu, "max_nsu": ultimo_nsu,
                "tem_mais": False, "xml_bruto": xml_bruto}


    def consultar_por_chave(self, cnpj: str, chave: str) -> dict:
        """
        consChNFe — NT 2014.002 seção 3.7 (v1.15).
        Consulta pontual de uma NF-e pela chave de acesso.
        NÃO precisa de NSU prévio — não consome o token de distNSU.
        Limite: 20 consultas por hora (NT 3.11.4.2).
        Retorna procNFe (XML completo) se o destinatário já manifestou.
        Retorna resNFe (resumo) se ainda não manifestou.
        """
        xml_body      = self._montar_xml_chave(cnpj, chave)
        soap_envelope = self._montar_soap(xml_body)
        ts            = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
        self._diag_log = []
        self._diag(f"consChNFe: chave={chave}")

        if self._modo_winhttp:
            return self._consultar_winhttp(soap_envelope, ts, 0)
        else:
            return self._consultar_requests(soap_envelope, ts, 0)

    def _montar_xml_chave(self, cnpj: str, chave: str) -> str:
        """consChNFe — NT seção 3.4.1c."""
        return (
            f'<distDFeInt xmlns="{NS_NFE}" versao="1.01">'
            f'<tpAmb>{self.ambiente}</tpAmb>'
            f'<cUFAutor>{self.cuf_autor}</cUFAutor>'
            f'<CNPJ>{cnpj}</CNPJ>'
            f'<consChNFe><chNFe>{chave}</chNFe></consChNFe>'
            f'</distDFeInt>'
        )
