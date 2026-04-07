"""
Storage local - persistência do NSU e das notas consultadas.
Usa arquivos JSON simples na pasta 'data/' do projeto.
O NSU é crítico: sem ele o SEFAZ rejeita a próxima consulta como "Consumo Indevido".
"""

import json
import os
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional

# Usa MPI_NFE_DIR quando rodando como .exe (PyInstaller), senão usa pasta do projeto
_BASE = Path(os.environ.get("MPI_NFE_DIR", Path(__file__).parent.parent))
DATA_DIR = _BASE / "data"


class StorageNSU:
    """
    Gerencia o estado persistente das consultas:
    - Último NSU por CNPJ
    - Notas recebidas por CNPJ
    """

    def __init__(self):
        DATA_DIR.mkdir(parents=True, exist_ok=True)
        self._nsu_file  = DATA_DIR / "nsu_estado.json"
        self._nota_dir  = DATA_DIR / "notas"
        self._nota_dir.mkdir(exist_ok=True)
        self._estado    = self._carregar_estado()

    def _carregar_estado(self) -> dict:
        if self._nsu_file.exists():
            try:
                with open(self._nsu_file, encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {}

    def _salvar_estado(self):
        with open(self._nsu_file, "w", encoding="utf-8") as f:
            json.dump(self._estado, f, indent=2, ensure_ascii=False)

    def salvar_cert_config(self, tipo: str, valor: str, cnpj: str = ""):
        """Salva configuração do certificado para restaurar na próxima abertura."""
        self._estado["_cert_config"] = {"tipo": tipo, "valor": valor, "cnpj": cnpj}
        self._salvar_estado()

    def get_cert_config(self) -> dict:
        """Retorna configuração salva do certificado."""
        return self._estado.get("_cert_config", {})

    def limpar_cert_config(self):
        self._estado.pop("_cert_config", None)
        self._salvar_estado()

    # ── NSU ──────────────────────────────────────────────────────────────────

    def get_nsu(self, cnpj: str) -> int:
        """Retorna o último NSU para o CNPJ. 0 se nunca consultado."""
        return int(self._estado.get(cnpj, {}).get("ultimo_nsu", 0))

    def set_nsu(self, cnpj: str, nsu: int):
        """Persiste o novo último NSU para o CNPJ."""
        if cnpj not in self._estado:
            self._estado[cnpj] = {}
        self._estado[cnpj]["ultimo_nsu"] = nsu
        self._estado[cnpj]["ultima_consulta"] = datetime.now().isoformat()
        self._salvar_estado()

    def ultima_consulta(self, cnpj: str) -> Optional[str]:
        return self._estado.get(cnpj, {}).get("ultima_consulta")

    # ── Notas ─────────────────────────────────────────────────────────────────

    def _nota_file(self, cnpj: str) -> Path:
        return self._nota_dir / f"notas_{cnpj}.json"

    def salvar_notas(self, cnpj: str, notas: List[Dict]):
        """Salva/atualiza notas. Novas são adicionadas; existentes são atualizadas."""
        existentes = self.get_todas_notas(cnpj)
        idx_por_chave = {n.get("chave", ""): i for i, n in enumerate(existentes)}

        for nota in notas:
            chave = nota.get("chave", "")
            if chave and chave in idx_por_chave:
                # Atualiza — preserva xml_raw se a nova não tiver ou for menor
                i = idx_por_chave[chave]
                xml_antigo = existentes[i].get("xml_raw", "")
                xml_novo   = nota.get("xml_raw", "")
                # Mantém o maior XML (procNFe tem mais chars que resNFe)
                if xml_antigo and (not xml_novo or len(xml_antigo) > len(xml_novo)):
                    nota["xml_raw"] = xml_antigo
                existentes[i] = nota
            else:
                existentes.append(nota)
                if chave:
                    idx_por_chave[chave] = len(existentes) - 1

        with open(self._nota_file(cnpj), "w", encoding="utf-8") as f:
            json.dump(existentes, f, indent=2, ensure_ascii=False)

    def get_todas_notas(self, cnpj: str) -> List[Dict]:
        """Retorna todas as notas salvas para o CNPJ."""
        path = self._nota_file(cnpj)
        if not cnpj or not path.exists():
            return []
        try:
            with open(path, encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []

    def get_nota(self, cnpj: str, chave: str) -> Optional[Dict]:
        """Retorna uma nota específica pela chave de acesso."""
        for nota in self.get_todas_notas(cnpj):
            if nota.get("chave", "").strip() == chave.strip():
                return nota
        return None

    def resumo(self, cnpj: str) -> dict:
        notas = self.get_todas_notas(cnpj)
        return {
            "cnpj": cnpj,
            "total_notas": len(notas),
            "ultimo_nsu": self.get_nsu(cnpj),
            "ultima_consulta": self.ultima_consulta(cnpj),
        }
