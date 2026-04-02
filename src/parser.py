"""
Parser de documentos fiscais retornados pelo SEFAZ (NFeDistribuicaoDFe).
Suporta: resNFe (resumo), procNFe (NF-e completa), procEventoNFe (evento/cancelamento).
"""

import xml.etree.ElementTree as ET
from datetime import datetime


NS = {
    "nfe":   "http://www.portalfiscal.inf.br/nfe",
    "sig":   "http://www.w3.org/2000/09/xmldsig#",
}

def _ns(tag: str) -> str:
    return f"{{{NS['nfe']}}}{tag}"


class ParserNFe:
    """Transforma XML retornado pelo SEFAZ em dict Python."""

    def parse_documento(self, xml_str: str, schema: str, nsu: int) -> dict:
        """
        :param xml_str: XML descomprimido (string UTF-8)
        :param schema:  valor do atributo 'schema' do docZip (ex: 'resNFe_v1.01.xsd')
        :param nsu:     número sequencial único deste documento
        """
        try:
            root = ET.fromstring(xml_str)
        except ET.ParseError:
            return {"nsu": nsu, "schema": schema, "chave": "", "situacao": "XML inválido"}

        schema_lower = schema.lower()

        if "resnfe" in schema_lower:
            return self._parse_resumo(root, nsu, schema)
        elif "proceventonfe" in schema_lower:
            return self._parse_evento(root, nsu, schema)
        elif "procnfe" in schema_lower or "nfe_v" in schema_lower:
            return self._parse_proc_nfe(root, nsu, schema)
        else:
            # Detecta pelo elemento raiz
            tag = root.tag.lower()
            if "resnfe" in tag:
                return self._parse_resumo(root, nsu, schema)
            elif "proceventonfe" in tag or "eventonfe" in tag:
                return self._parse_evento(root, nsu, schema)
            else:
                return self._parse_proc_nfe(root, nsu, schema)

    # ── Resumo NF-e (resNFe) ─────────────────────────────────────────────────

    def _parse_resumo(self, root: ET.Element, nsu: int, schema: str) -> dict:
        def t(tag):
            el = root.find(f".//{_ns(tag)}")
            return el.text.strip() if el is not None and el.text else ""

        chave   = t("chNFe")
        emissao = self._fmt_data(t("dhEmi"))
        valor   = self._fmt_valor(t("vNF"))
        cnpj_e  = t("CNPJ")
        xnome   = t("xNome")
        cstat   = t("cSitNFe")
        sit     = self._situacao_nfe(cstat)

        return {
            "nsu":      nsu,
            "schema":   schema,
            "tipo":     "Resumo",
            "chave":    chave,
            "emitente": f"{xnome} ({cnpj_e})" if xnome else cnpj_e,
            "valor":    valor,
            "emissao":  emissao,
            "situacao": sit,
            "cstat":    cstat,
            "xml_raw":  "",  # resumo não traz XML completo
        }

    # ── NF-e Completa (procNFe) ───────────────────────────────────────────────

    def _parse_proc_nfe(self, root: ET.Element, nsu: int, schema: str) -> dict:
        def t(tag):
            el = root.find(f".//{_ns(tag)}")
            return el.text.strip() if el is not None and el.text else ""

        chave   = t("chNFe") or self._chave_do_id(root)
        emissao = self._fmt_data(t("dhEmi"))
        valor   = self._fmt_valor(t("vNF"))
        cnpj_e  = t("CNPJ")
        xnome   = t("xNome")
        cstat_p = t("cStat")  # do protNFe
        sit     = self._situacao_nfe(cstat_p or "100")

        return {
            "nsu":      nsu,
            "schema":   schema,
            "tipo":     "NF-e Completa",
            "chave":    chave,
            "emitente": f"{xnome} ({cnpj_e})" if xnome else cnpj_e,
            "valor":    valor,
            "emissao":  emissao,
            "situacao": sit,
            "cstat":    cstat_p,
            "xml_raw":  "",
        }

    # ── Evento NF-e (procEventoNFe) ──────────────────────────────────────────

    def _parse_evento(self, root: ET.Element, nsu: int, schema: str) -> dict:
        def t(tag):
            el = root.find(f".//{_ns(tag)}")
            return el.text.strip() if el is not None and el.text else ""

        chave   = t("chNFe")
        tp_ev   = t("tpEvento")
        xdesc   = t("xEvento") or t("xCondUso") or self._desc_evento(tp_ev)
        dhevt   = self._fmt_data(t("dhEvento") or t("dhRegEvento"))

        return {
            "nsu":      nsu,
            "schema":   schema,
            "tipo":     f"Evento ({xdesc})",
            "chave":    chave,
            "emitente": t("CNPJ"),
            "valor":    "—",
            "emissao":  dhevt,
            "situacao": xdesc or "Evento",
            "cstat":    t("cStat"),
            "xml_raw":  "",
        }

    # ── Helpers ───────────────────────────────────────────────────────────────

    @staticmethod
    def _fmt_data(raw: str) -> str:
        if not raw:
            return ""
        for fmt in ("%Y-%m-%dT%H:%M:%S%z", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d"):
            try:
                dt = datetime.strptime(raw[:19], fmt[:len(raw[:19])])
                return dt.strftime("%d/%m/%Y %H:%M")
            except ValueError:
                pass
        return raw[:16]

    @staticmethod
    def _fmt_valor(raw: str) -> str:
        if not raw:
            return ""
        try:
            return f"{float(raw):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except ValueError:
            return raw

    @staticmethod
    def _situacao_nfe(cstat: str) -> str:
        mapa = {
            # cSitNFe do resNFe
            "1": "Autorizada",
            "2": "Cancelada",
            "3": "Denegada",
            # cStat do protNFe / evento
            "100": "Autorizada",
            "101": "Cancelada",
            "102": "Inutilizada",
            "110": "Denegada",
            "150": "Autorizada",
            "135": "Cancelada",
        }
        return mapa.get(cstat, f"Código {cstat}" if cstat else "Desconhecida")

    @staticmethod
    def _desc_evento(tp_ev: str) -> str:
        mapa = {
            "110111": "Cancelamento",
            "110110": "Carta de Correção",
            "110140": "EPEC",
            "210200": "Confirmação da Operação",
            "210210": "Ciência da Operação",
            "210220": "Desconhecimento da Operação",
            "210240": "Operação não Realizada",
        }
        return mapa.get(tp_ev, f"Evento {tp_ev}")

    @staticmethod
    def _chave_do_id(root: ET.Element) -> str:
        """Extrai a chave de acesso do atributo Id da infNFe."""
        inf = root.find(f".//{_ns('infNFe')}")
        if inf is not None:
            id_val = inf.get("Id", "")
            if id_val.startswith("NFe"):
                return id_val[3:]
        return ""
