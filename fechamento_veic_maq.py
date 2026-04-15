"""
Robô Fechamento Veic/Maq - MultiA
===================================
Automação de cadastro de comparativos de veículos/máquinas nos sistemas
MultiA Mais e MultiA Avaliações.

Lê dados de planilhas Google Sheets e envia comparativos via API REST.
100% em segundo plano, sem navegador.

Campos da planilha (linha encontrada em K16:K25 pelo número da imagem):
  K = número do comparativo
  L = fonte
  M = ano/modelo  (→ ANOMODELO)
  N = KM/horas   (→ KM)
  O = valor      (→ VALOR)
  R = OBS        (→ ignorado se vazio)

Campos fixos da planilha (lidos uma vez por aba):
  B4  = Modelo do bem        (→ MODELO, usado em Máquinas)
  D5  = Obs. Laudo           (→ OBSLAUDO)
  B13 = Tipo do bem          (usado para detectar Máquina)
  B29 = Valor FIPE           (→ VLRFIPE)
  B31 = Referência considerada (→ REFCONSID)
  B34 = Fator Depreciação    (→ FATORDEP)  | D34 = motivo
  B35 = Fator Valorização    (→ FATORVALO) | D35 = motivo (ignorado se vazio)
  B36 = Fator Comercialização (→ FATORCOMERC, invertido) | D36 = motivo
  B40 = Liq. Forçada         (→ PERCENTFORCADA)

Uso:
    python fechamento_veic_maq.py

Compilar:
    pyinstaller --onefile --windowed --icon="logo-fechamento-carros.ico" fechamento_veic_maq.py
"""

import os
import sys
import re
import json
import time
import math
import logging
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
from pathlib import Path
from threading import Thread
from dataclasses import dataclass
from typing import Optional

import customtkinter as ctk
import requests
import gspread
from google.oauth2.service_account import Credentials


# ============================================================
# UTILITÁRIO
# ============================================================

def _base_dir() -> Path:
    """Retorna a pasta do EXE (compilado) ou do script (dev)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).parent


# ============================================================
# CONSTANTES
# ============================================================

def _carregar_sistema(nome: str) -> dict:
    """Carrega configuração do sistema a partir de variáveis de ambiente."""
    prefixo = nome.upper().replace(" ", "_").replace("Ã", "A").replace("Ç", "C")
    return {
        "base_url":      os.environ.get(f"{prefixo}_BASE_URL", ""),
        "origin":        os.environ.get(f"{prefixo}_ORIGIN", ""),
        "referer":       os.environ.get(f"{prefixo}_REFERER", ""),
        "authorization": os.environ.get(f"{prefixo}_AUTHORIZATION", ""),
        "jwt_fixo":      os.environ.get(f"{prefixo}_JWT", ""),
    }

SISTEMAS = {
    "MultiA Mais":       _carregar_sistema("MULTIA_MAIS"),
    "MultiA Avaliações": _carregar_sistema("MULTIA_AVALIACOES"),
}

IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp"}

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Linhas da planilha onde ficam os comparativos (K16:K25 → índices 15..24)
COMP_ROW_START = 15
COMP_ROW_END   = 25

# Colunas (índices base-0)
COL_K     = 10  # número do comparativo
COL_L     = 11  # fonte
COL_M     = 12  # ano/modelo
COL_N     = 13  # KM / horas
COL_O     = 14  # valor
COL_J     = 17  # OBS do comparativo (coluna R)

# Células fixas por aba
ROW_VALORFIPE  = 28   # B29
ROW_REFCONSID  = 30   # B31
ROW_FATORDEP   = 33   # B34
ROW_FATORVALOR = 34   # B35
ROW_FATORCOMERC= 35   # B36
ROW_LIQFORCADA = 39   # B40
ROW_OBSLAUDO   = 4    # D5
ROW_TIPO       = 12   # B13
ROW_MODELO     = 3    # B4


# ============================================================
# DATACLASSES
# ============================================================

@dataclass
class ComparativoVeicData:
    numero:      int
    fonte:       str
    ano_modelo:  str
    km:          str
    valor:       str
    obs:         str
    imagem_path: str


@dataclass
class ConfigData:
    sistema:           str  = "MultiA Mais"
    planilha_id:       str  = ""
    pasta_comparativos:str  = ""
    credentials_path:  str  = ""
    excluir_imagens:   bool = True
    validade_laudo:    str  = "12"


# ============================================================
# CLIENTE API MULTIA
# ============================================================

class MultiAAPI:
    """Acesso à API MultiA via HTTP puro — sem navegador."""

    def __init__(self, sistema_cfg: dict, logger: logging.Logger):
        self.cfg = sistema_cfg
        self.logger = logger
        self.base_url = sistema_cfg["base_url"]
        self.session = requests.Session()
        self.session.headers.update({
            "authorization": sistema_cfg["authorization"],
            "jwt":           sistema_cfg["jwt_fixo"],
            "Accept":        "*/*",
            "Origin":        sistema_cfg["origin"],
            "Referer":       sistema_cfg["referer"],
        })

    def buscar_avaliacoes(self, busca: str, page_size: int = 50) -> dict:
        url = f"{self.base_url}/multia/avaliacoes"
        params = {
            "sortField": "", "sortOrder": "",
            "pageSize": page_size, "page": 0,
            "busca": busca,
            "REGSTATUS": "", "DATACRIACAO": "",
        }
        self.logger.info(f"  GET {url} busca={busca}")
        resp = self.session.get(url, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def buscar_avaliacao_por_codigo(self, codigo: str) -> Optional[dict]:
        """Retorna o dict da avaliação cujo REG coincide com `codigo`, ou None."""
        data = self.buscar_avaliacoes(codigo)
        if data.get("status") != "sucesso":
            return None
        for av in data.get("dados", {}).get("avaliacoes", []):
            if str(av.get("REG", "")).strip() == str(codigo).strip():
                return av
        return None

    def buscar_dados_avaliacao(self, uuid: str) -> dict:
        url = f"{self.base_url}/multia/dadosavaliacao/{uuid}"
        resp = self.session.get(url, timeout=30)
        resp.raise_for_status()
        return resp.json()

    def adicionar_comparativo(self, uuid: str,
                               ano_modelo: str, km: str,
                               valor: str, fonte: str,
                               obs: str,
                               imagem_path: str,
                               modelo: str = "",
                               is_maquina: bool = False) -> dict:
        """
        Envia POST /multia/adicionarcomparativo/{uuid} com multipart/form-data.
        Para Máquinas, envia FABRICACAO em vez de ANOMODELO, e inclui MODELO.
        OBS só é enviado se não estiver vazio.
        """
        url = f"{self.base_url}/multia/adicionarcomparativo/{uuid}"

        data: dict = {
            ("FABRICACAO" if is_maquina else "ANOMODELO"): ano_modelo,
            "KM":        km,
            "HORASUSO":  km,
            "VALOR":     valor,
            "FONTE":     fonte,
        }
        if modelo:
            data["MODELO"] = modelo
        if obs:
            data["OBSERVACOES"] = obs

        filename = os.path.basename(imagem_path)
        ext = Path(imagem_path).suffix.lower()
        mime = "image/png" if ext == ".png" else "image/jpeg"

        with open(imagem_path, "rb") as f:
            files = {"arquivo": (filename, f, mime)}
            resp = self.session.post(url, data=data, files=files, timeout=60)

        resp.raise_for_status()
        return resp.json()

    def editar_avaliacao(self, uuid: str, **campos) -> dict:
        url = f"{self.base_url}/multia/editaravaliacao/{uuid}"
        self.logger.info(f"  POST {url} → {list(campos.keys())}")
        resp = self.session.post(url, data=campos, timeout=30)
        resp.raise_for_status()
        return resp.json()


# ============================================================
# CLIENTE GOOGLE SHEETS
# ============================================================

class PlanilhaClient:
    """Lê dados da planilha de comparativos de veículos/máquinas."""

    def __init__(self, credentials_path: str, planilha_id: str,
                 logger: logging.Logger):
        self.logger = logger
        creds = Credentials.from_service_account_file(credentials_path, scopes=SCOPES)
        self.gc = gspread.authorize(creds)
        self.planilha = self.gc.open_by_key(planilha_id)
        self.logger.info(f"  Planilha conectada: {self.planilha.title}")

    def ler_dados_subpasta(self, nome_subpasta: str,
                            numeros_imagens: list[int]
                            ) -> tuple:
        """
        Abre a aba com o mesmo nome da subpasta e lê os dados dos comparativos
        e as células fixas de fatores e percentuais.
        """
        self.logger.info(f"  Abrindo aba '{nome_subpasta}'...")

        try:
            aba = self.planilha.worksheet(nome_subpasta)
        except gspread.exceptions.WorksheetNotFound:
            alt = re.sub(r'[\.\-]', '', nome_subpasta)
            try:
                aba = self.planilha.worksheet(alt)
                self.logger.warning(f"  Aba '{nome_subpasta}' não encontrada — usando '{alt}'")
            except gspread.exceptions.WorksheetNotFound:
                raise ValueError(f"Aba '{nome_subpasta}' não encontrada na planilha")

        todos = aba.get_all_values()

        comparativos: list[ComparativoVeicData] = []

        for num_img in sorted(numeros_imagens):
            encontrado = False
            for row_idx in range(COMP_ROW_START, min(COMP_ROW_END, len(todos))):
                row = todos[row_idx]
                if not row:
                    continue

                celula_k = row[COL_K].strip() if len(row) > COL_K else ""
                try:
                    num_celula = int(re.sub(r'\D', '', celula_k)) if celula_k else -1
                except ValueError:
                    continue

                if num_celula == num_img:
                    def _get(col: int) -> str:
                        return row[col].strip() if len(row) > col else ""

                    comparativos.append(ComparativoVeicData(
                        numero=num_img,
                        fonte=_get(COL_L),
                        ano_modelo=_get(COL_M),
                        km=_get(COL_N),
                        valor=_get(COL_O),
                        obs=_get(COL_J),
                        imagem_path="",
                    ))
                    encontrado = True
                    break

            if not encontrado:
                self.logger.warning(
                    f"    Imagem {num_img}: linha não encontrada em K16:K25"
                )

        def _cell(row_idx: int, col_idx: int) -> str:
            if len(todos) > row_idx and len(todos[row_idx]) > col_idx:
                return todos[row_idx][col_idx].strip()
            return ""

        valor_fipe   = _cell(ROW_VALORFIPE,   1)
        ref_consid   = _cell(ROW_REFCONSID,   1)
        fator_dep    = _cell(ROW_FATORDEP,    1)
        motiv_dep    = _cell(ROW_FATORDEP,    3)
        fator_valor  = _cell(ROW_FATORVALOR,  1)
        motiv_valor  = _cell(ROW_FATORVALOR,  3)
        fator_comerc = _cell(ROW_FATORCOMERC, 1)
        motiv_comerc = _cell(ROW_FATORCOMERC, 3)
        liq_forcada  = _cell(ROW_LIQFORCADA,  1)
        obs_laudo    = _cell(ROW_OBSLAUDO,    3)
        tipo_bem     = _cell(ROW_TIPO,        1)
        modelo_bem   = _cell(ROW_MODELO,      1)

        return (comparativos, valor_fipe, ref_consid, fator_dep, motiv_dep,
                fator_valor, motiv_valor, fator_comerc, motiv_comerc,
                liq_forcada, obs_laudo, tipo_bem, modelo_bem)


# ============================================================
# MOTOR DO ROBÔ
# ============================================================

class RoboFechamento:
    """Orquestra o processo de fechamento de veíc/máq em background."""

    def __init__(self, config: ConfigData, logger: logging.Logger,
                 callback_progresso=None, callback_confirmar=None,
                 callback_validade=None):
        self.config             = config
        self.logger             = logger
        self.callback_progresso = callback_progresso or (lambda m: None)
        self.callback_confirmar = callback_confirmar
        self.callback_validade  = callback_validade
        self.api:      Optional[MultiAAPI]      = None
        self.planilha: Optional[PlanilhaClient] = None
        self._cancelado = False

    def cancelar(self):
        self._cancelado = True

    def _log(self, msg: str):
        self.logger.info(msg)
        self.callback_progresso(msg)

    def _listar_subpastas(self) -> list[tuple[str, list[str]]]:
        pasta = Path(self.config.pasta_comparativos)
        resultado = []
        for sub in sorted(pasta.iterdir()):
            if not sub.is_dir():
                continue
            imagens = sorted(
                [f.name for f in sub.iterdir()
                 if f.suffix.lower() in IMAGE_EXTENSIONS],
                key=lambda x: int(re.sub(r'\D', '', Path(x).stem) or 0)
            )
            if imagens:
                resultado.append((sub.name, imagens))
        return resultado

    def _processar_subpasta(self, nome: str, imagens: list[str]):
        self._log(f"\n{'='*60}")
        self._log(f"PROCESSANDO: {nome}")
        self._log(f"{'='*60}")

        pasta_imgs = Path(self.config.pasta_comparativos) / nome

        mapa_imagens: dict[int, str] = {}
        for img in imagens:
            stem = Path(img).stem
            try:
                num = int(re.sub(r'\D', '', stem))
                mapa_imagens[num] = str(pasta_imgs / img)
            except ValueError:
                self._log(f"  AVISO: '{img}' ignorada (nome não numérico)")

        if not mapa_imagens:
            self._log("  Nenhuma imagem com nome numérico encontrada")
            return

        numeros = sorted(mapa_imagens.keys())
        self._log(f"  Imagens encontradas: {numeros}")

        try:
            (comparativos, valor_fipe, ref_consid, fator_dep, motiv_dep,
             fator_valor, motiv_valor, fator_comerc, motiv_comerc,
             liq_forcada, obs_laudo, tipo_bem, modelo_bem) = \
                self.planilha.ler_dados_subpasta(nome, numeros)
        except Exception as e:
            self._log(f"  ERRO ao ler planilha: {e}")
            return

        for comp in comparativos:
            comp.imagem_path = mapa_imagens.get(comp.numero, "")

        self._log(f"  {len(comparativos)} comparativo(s) lido(s) da planilha")
        if obs_laudo:    self._log(f"  Obs. Laudo (D5):           {obs_laudo}")
        if valor_fipe:   self._log(f"  Valor FIPE (B29):          {valor_fipe}")
        if ref_consid:   self._log(f"  Ref. considerada (B31):    {ref_consid}")
        if fator_dep:    self._log(f"  Fator Depreciação (B34):   {fator_dep} | motivo: {motiv_dep}")
        if fator_valor:  self._log(f"  Fator Valorização (B35):   {fator_valor} | motivo: {motiv_valor}")
        else:            self._log(f"  Fator Valorização (B35):   (vazio — ignorado)")
        if fator_comerc: self._log(f"  Fator Comercializ. (B36):  {fator_comerc} | motivo: {motiv_comerc}")
        if liq_forcada:  self._log(f"  Liq. Forçada (B40):        {liq_forcada}")

        is_maquina = "máquina" in tipo_bem.lower() or "maquina" in tipo_bem.lower()
        if is_maquina:
            self._log(f"  ★ Tipo MÁQUINA → MODELO (B4): '{modelo_bem or '(vazio!)'}'")

        # Buscar avaliação
        self._log(f"  Buscando avaliação: '{nome}' | Tipo: '{tipo_bem}'")
        avaliacao = None
        try:
            result = self.api.buscar_avaliacoes(nome)
            lista = result.get("dados", {}).get("avaliacoes", [])

            nome_lower = nome.strip().lower()
            tipo_lower = tipo_bem.strip().lower()

            nao_imovel = [
                av for av in lista
                if "imóvel" not in (av.get("PRODUTO") or "").lower()
                and "imovel" not in (av.get("PRODUTO") or "").lower()
            ]

            for av in nao_imovel:
                doc     = str(av.get("documento") or "").strip().lower()
                produto = str(av.get("PRODUTO") or "").strip().lower()
                if doc == nome_lower and tipo_lower in produto:
                    avaliacao = av
                    self._log(
                        f"  ✓ Match: REG={av.get('REG')} | "
                        f"Doc={av.get('documento')} | Produto={av.get('PRODUTO')}"
                    )
                    break

            if not avaliacao:
                for av in nao_imovel:
                    doc = str(av.get("documento") or "").strip().lower()
                    if doc == nome_lower:
                        avaliacao = av
                        self._log(
                            f"  ⚠ Match parcial (tipo divergente): "
                            f"REG={av.get('REG')} | Produto={av.get('PRODUTO')}"
                        )
                        break

            if not avaliacao:
                self._log(f"  ✗ Avaliação não encontrada para '{nome}'")
                for av in lista:
                    self._log(
                        f"    REG={av.get('REG')} | Doc={av.get('documento')} "
                        f"| Produto={av.get('PRODUTO')} | Status={av.get('STATUS','')}"
                    )

        except Exception as e:
            self._log(f"  ERRO na busca: {e}")

        if not avaliacao:
            self._log(f"  ✗ Pulando '{nome}'")
            return

        uuid = avaliacao.get("UUID") or avaliacao.get("uuid")
        if not uuid:
            self._log("  ERRO: UUID não encontrado — pulando")
            return

        self._log(f"  ✓ UUID: {uuid}  |  Status: {avaliacao.get('STATUS','')}")
        self.callback_progresso(f"__UUID__:{uuid}")

        # Cadastrar comparativos
        total_comp = len(comparativos)
        for i, comp in enumerate(comparativos, 1):
            if self._cancelado:
                self._log("  CANCELADO pelo usuário")
                return

            if not comp.imagem_path or not os.path.exists(comp.imagem_path):
                self._log(f"  [{i}/{total_comp}] Comp {comp.numero}: imagem não encontrada")
                continue

            self._log(
                f"  [{i}/{total_comp}] Comp {comp.numero}: "
                f"ano={comp.ano_modelo} | km={comp.km} | valor={comp.valor}"
            )

            try:
                resultado = self.api.adicionar_comparativo(
                    uuid=uuid,
                    ano_modelo=comp.ano_modelo,
                    km=comp.km,
                    valor=comp.valor,
                    fonte=comp.fonte,
                    obs=comp.obs,
                    imagem_path=comp.imagem_path,
                    modelo=modelo_bem if is_maquina else "",
                    is_maquina=is_maquina,
                )

                if resultado.get("status") == "sucesso":
                    self._log(f"    ✓ Adicionado (REG: {resultado.get('dados')})")
                    if self.config.excluir_imagens:
                        try:
                            os.remove(comp.imagem_path)
                        except OSError as e:
                            self._log(f"    AVISO: não foi possível excluir imagem: {e}")
                else:
                    self._log(f"    ✗ Resposta inesperada: {resultado}")

            except requests.exceptions.HTTPError as e:
                self._log(f"    ✗ Erro HTTP {e.response.status_code}: {e.response.text[:200]}")
            except Exception as e:
                self._log(f"    ✗ Erro: {e}")

        # Ativar toggle Avaliação
        self._log(f"\n  [1/2] Ativando toggle Avaliação (MEMORIALCALC=S)...")
        try:
            res = self.api.editar_avaliacao(uuid, MEMORIALCALC="S")
            if res.get("status") == "sucesso":
                self._log("    ✓ Toggle ativado")
            else:
                self._log(f"    ✗ Erro: {res}")
        except Exception as e:
            self._log(f"    ✗ Erro ao ativar toggle: {e}")

        # Enviar campos da avaliação
        validade = (self.callback_validade() if self.callback_validade else "").strip()
        campos_editar: dict = {"PERCENTJUSTA": "10"}

        if obs_laudo:
            campos_editar["OBSLAUDO"] = obs_laudo
        if valor_fipe:
            campos_editar["VLRFIPE"] = re.sub(r'[^\d]', '', valor_fipe.split(',')[0])
        if ref_consid:
            campos_editar["REFCONSID"] = ref_consid
        if fator_dep:
            campos_editar["FATORDEP"] = re.sub(r'[^\d\-\.]', '', fator_dep.replace(',', '.'))
            if motiv_dep:
                campos_editar["MOTIVFATORDEP"] = motiv_dep
        if fator_valor:
            campos_editar["FATORVALO"] = re.sub(r'[^\d\-\.]', '', fator_valor.replace(',', '.'))
            if motiv_valor:
                campos_editar["MOTIVFATORVALO"] = motiv_valor
        if fator_comerc:
            # Fator de comercialização: invertido em relação à planilha
            fc_num = re.sub(r'[^\d\-\.]', '', fator_comerc.replace(',', '.'))
            try:
                fc_inv = str(-float(fc_num))
                if fc_inv.endswith('.0'):
                    fc_inv = fc_inv[:-2]
            except ValueError:
                fc_inv = fc_num
            campos_editar["FATORCOMERC"] = fc_inv
            if motiv_comerc:
                campos_editar["MOTIVFATORCOMERC"] = motiv_comerc
        if liq_forcada:
            campos_editar["PERCENTFORCADA"] = re.sub(r'[^\d\-\.]', '', liq_forcada.replace(',', '.'))
        if validade:
            campos_editar["VALIDADELAUDO"] = validade

        if campos_editar:
            self._log(f"  [2/2] Enviando campos: {list(campos_editar.keys())}...")
            try:
                res = self.api.editar_avaliacao(uuid, **campos_editar)
                if res.get("status") == "sucesso":
                    self._log("    ✓ Campos enviados com sucesso")
                else:
                    self._log(f"    ✗ Erro: {res}")
            except requests.exceptions.HTTPError as e:
                self._log(f"    ✗ Erro HTTP {e.response.status_code}: {e.response.text[:200]}")
            except Exception as e:
                self._log(f"    ✗ Erro: {e}")

        self._log(f"\n  '{nome}' finalizado!")

    def executar(self):
        try:
            self._log("=" * 60)
            self._log("ROBÔ FECHAMENTO VEIC/MÁQ - INICIANDO")
            self._log("=" * 60)

            subpastas = self._listar_subpastas()
            if not subpastas:
                self._log("Nenhuma subpasta com imagens encontrada!")
                return

            self._log(f"\nSubpastas encontradas:")
            for nome, imgs in subpastas:
                self._log(f"  📁 {nome}  ({len(imgs)} imagem(ns))")

            self._log(f"\nConectando ao Google Sheets...")
            try:
                self.planilha = PlanilhaClient(
                    self.config.credentials_path,
                    self.config.planilha_id,
                    self.logger,
                )
            except Exception as e:
                self._log(f"ERRO ao conectar planilha: {e}")
                return

            sistema_cfg = SISTEMAS[self.config.sistema]
            self._log(f"\nConectando API ({self.config.sistema})...")
            self.api = MultiAAPI(sistema_cfg, self.logger)

            try:
                teste = self.api.buscar_avaliacoes("9999999", page_size=1)
                if teste.get("status") == "sucesso":
                    self._log("  ✓ API conectada com sucesso!")
                else:
                    self._log(f"  ✗ API retornou: {teste}")
                    return
            except Exception as e:
                self._log(f"  ✗ Falha na conexão: {e}")
                return

            total = len(subpastas)
            for idx, (nome, imgs) in enumerate(subpastas, 1):
                if self._cancelado:
                    self._log("\nEXECUÇÃO CANCELADA")
                    break
                self.callback_progresso(f"PROGRESSO:{idx}/{total}")
                self._processar_subpasta(nome, imgs)

            self._log("\n" + "=" * 60)
            self._log("EXECUÇÃO FINALIZADA")
            self._log("=" * 60)

        except Exception as e:
            self._log(f"\nERRO FATAL: {e}")
            import traceback
            self._log(traceback.format_exc())


# ============================================================
# INTERFACE GRÁFICA
# ============================================================

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

C = {
    "bg":         "#0B0F1A",
    "surface":    "#111827",
    "surface2":   "#1A2236",
    "border":     "#1E2D45",
    "text":       "#F0F4FF",
    "muted":      "#6B7FA3",
    "accent":     "#F5C400",
    "accent_hov": "#D4A900",
    "accent_dim": "#2A2400",
    "blue":       "#1D4ED8",
    "blue_hov":   "#1E40AF",
    "blue_lite":  "#1E3A5F",
    "neutral":    "#6B7FA3",
    "neut_bg":    "#1A2236",
    "neut_brd":   "#1E2D45",
    "ok":         "#22C55E",
    "err":        "#EF4444",
    "warn":       "#F59E0B",
    "log_bg":     "#070B14",
    "log_fg":     "#CBD5E1",
    "log_ok":     "#4ADE80",
    "log_err":    "#F87171",
    "log_warn":   "#FBBF24",
    "log_info":   "#60A5FA",
    "log_muted":  "#334155",
}

FONT_BTN   = ("Segoe UI", 13, "bold")
FONT_INPUT = ("Segoe UI", 12)
FONT_LOG   = ("Consolas", 10)
FONT_MONO  = ("Cascadia Code", 10)


class App:
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("MultiA — Fechamento Veic/Maq")
        self.root.geometry("860x660")
        self.root.resizable(True, True)
        self.root.minsize(760, 580)
        self.root.configure(fg_color=C["bg"])
        self.root.configure(bg=C["bg"])

        self.config     = ConfigData()
        self.robo:      Optional[RoboFechamento] = None
        self.api:       Optional[MultiAAPI]      = None
        self.executando = False

        self._ultimo_uuid:       str = ""
        self._validade_pendente: str = ""

        self._setup_logger()
        self._build_ui()
        self._carregar_config()

    def _setup_logger(self):
        self.logger = logging.getLogger("FechamentoVeicMaq")
        self.logger.setLevel(logging.DEBUG)
        if not self.logger.handlers:
            h = logging.StreamHandler(sys.stdout)
            h.setLevel(logging.DEBUG)
            h.setFormatter(logging.Formatter("%(asctime)s - %(message)s", datefmt="%H:%M:%S"))
            self.logger.addHandler(h)

    def _build_ui(self):
        root = self.root

        header = ctk.CTkFrame(root, fg_color=C["surface"],
                               corner_radius=0, height=48, border_width=0)
        header.pack(fill="x")
        header.pack_propagate(False)

        title_frame = ctk.CTkFrame(header, fg_color="transparent")
        title_frame.pack(side="left", padx=18, pady=8)
        ctk.CTkLabel(title_frame, text="MultiA",
                     font=("Segoe UI", 15, "bold"),
                     text_color=C["accent"]).pack(side="left")
        ctk.CTkLabel(title_frame, text="  Fechamento Veic/Maq",
                     font=("Segoe UI", 11),
                     text_color=C["muted"]).pack(side="left")

        self._dot_canvas = tk.Canvas(header, width=10, height=10,
                                      bg=C["surface"], highlightthickness=0)
        self._dot_canvas.pack(side="right", padx=18, pady=19)
        self._dot = self._dot_canvas.create_oval(1, 1, 9, 9, fill=C["muted"], outline="")

        ctk.CTkFrame(root, fg_color=C["border"], height=1, corner_radius=0).pack(fill="x")

        body = ctk.CTkFrame(root, fg_color=C["bg"])
        body.pack(fill="both", expand=True, padx=20, pady=12)
        body.columnconfigure(0, weight=2, minsize=320)
        body.columnconfigure(1, weight=3, minsize=360)
        body.rowconfigure(0, weight=1)

        left = ctk.CTkFrame(body, fg_color="transparent")
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

        self._section_title(left, "Sistema")
        card_sis = self._card(left)
        self.var_sistema = ctk.StringVar(value="MultiA Mais")
        radio_row = ctk.CTkFrame(card_sis, fg_color="transparent")
        radio_row.pack(fill="x")
        for nome in SISTEMAS:
            ctk.CTkRadioButton(
                radio_row, text=nome, variable=self.var_sistema, value=nome,
                font=("Segoe UI", 13), text_color=C["text"],
                fg_color=C["accent"], hover_color=C["accent_hov"],
                border_color=C["border"],
            ).pack(side="left", padx=(0, 20))

        self._section_title(left, "ID da Planilha Google Sheets")
        card_pl = self._card(left)
        self.entry_planilha = ctk.CTkEntry(
            card_pl, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="Cole o ID da planilha...",
            placeholder_text_color=C["muted"],
        )
        self.entry_planilha.pack(fill="x")

        self._section_title(left, "Credenciais Google")
        card_cr = self._card(left)
        row_cr = ctk.CTkFrame(card_cr, fg_color="transparent")
        row_cr.pack(fill="x")
        self.entry_creds = ctk.CTkEntry(
            row_cr, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="credentials.json",
            placeholder_text_color=C["muted"],
        )
        self.entry_creds.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            row_cr, text="📂", font=("Segoe UI", 10),
            fg_color=C["neut_bg"], text_color=C["text"],
            hover_color=C["blue_lite"], border_width=1, border_color=C["neut_brd"],
            corner_radius=8, height=44, width=54,
            command=self._selecionar_credentials,
        ).pack(side="left")

        self._section_title(left, "Pasta Comparativos")
        card_pa = self._card(left)
        row_pa = ctk.CTkFrame(card_pa, fg_color="transparent")
        row_pa.pack(fill="x")
        self.entry_pasta = ctk.CTkEntry(
            row_pa, font=FONT_INPUT,
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["text"], border_width=1, corner_radius=8, height=44,
            placeholder_text="Selecione a pasta...",
            placeholder_text_color=C["muted"],
        )
        self.entry_pasta.pack(side="left", fill="x", expand=True, padx=(0, 8))
        ctk.CTkButton(
            row_pa, text="📂", font=("Segoe UI", 10),
            fg_color=C["neut_bg"], text_color=C["text"],
            hover_color=C["blue_lite"], border_width=1, border_color=C["neut_brd"],
            corner_radius=8, height=44, width=54,
            command=self._selecionar_pasta,
        ).pack(side="left")

        self.lbl_subpastas = ctk.CTkLabel(
            card_pa, text="Nenhuma pasta selecionada",
            font=("Segoe UI", 11), text_color=C["muted"], anchor="w",
        )
        self.lbl_subpastas.pack(anchor="w", pady=(6, 0))

        opt_row = ctk.CTkFrame(left, fg_color="transparent")
        opt_row.pack(fill="x", pady=(8, 0))
        self.var_excluir = ctk.BooleanVar(value=True)
        ctk.CTkCheckBox(
            opt_row, text="Excluir imagens após envio",
            variable=self.var_excluir,
            font=("Segoe UI", 12), text_color=C["muted"],
            fg_color=C["accent"], hover_color=C["accent_hov"],
            border_color=C["border"], checkmark_color="#000000",
        ).pack(side="left")

        btn_row = ctk.CTkFrame(left, fg_color="transparent")
        btn_row.pack(fill="x", pady=(10, 0))

        self.btn_executar = ctk.CTkButton(
            btn_row, text="▶  EXECUTAR",
            font=FONT_BTN, fg_color=C["accent"], hover_color=C["accent_hov"],
            text_color="#0B0F1A", corner_radius=10, height=52,
            command=self._executar,
        )
        self.btn_executar.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.btn_cancelar = ctk.CTkButton(
            btn_row, text="■  PARAR",
            font=FONT_BTN, fg_color=C["neut_bg"], hover_color=C["blue_lite"],
            text_color=C["muted"], border_width=1, border_color=C["neut_brd"],
            corner_radius=10, height=52, state="disabled",
            command=self._cancelar,
        )
        self.btn_cancelar.pack(side="left", fill="x", expand=True)

        status_card = ctk.CTkFrame(left, fg_color=C["surface"],
                                    border_color=C["border"], border_width=1,
                                    corner_radius=8)
        status_card.pack(fill="x", pady=(10, 0))
        s_inner = ctk.CTkFrame(status_card, fg_color="transparent")
        s_inner.pack(fill="x", padx=12, pady=8)
        ctk.CTkLabel(s_inner, text="STATUS",
                     font=("Segoe UI", 7, "bold"),
                     text_color=C["muted"]).pack(side="left")
        ctk.CTkFrame(s_inner, fg_color=C["border"], width=1, height=14,
                     corner_radius=0).pack(side="left", padx=10)
        self.progress_var = ctk.StringVar(value="Aguardando...")
        ctk.CTkLabel(s_inner, textvariable=self.progress_var,
                     font=("Segoe UI", 9, "bold"),
                     text_color=C["accent"], anchor="w",
                     ).pack(side="left", fill="x", expand=True)

        right = ctk.CTkFrame(body, fg_color="transparent")
        right.grid(row=0, column=1, sticky="nsew")

        self._section_title(right, "Log de Execução")
        log_card = ctk.CTkFrame(right, fg_color=C["log_bg"],
                                 corner_radius=10, border_color=C["border"],
                                 border_width=1)
        log_card.pack(fill="both", expand=True)

        log_hdr = ctk.CTkFrame(log_card, fg_color=C["surface"],
                                corner_radius=0, height=28)
        log_hdr.pack(fill="x")
        log_hdr.pack_propagate(False)
        dots = ctk.CTkFrame(log_hdr, fg_color="transparent")
        dots.pack(side="left", padx=10, pady=7)
        for col in ("#FF5F57", "#FFBD2E", "#28C840"):
            ctk.CTkFrame(dots, fg_color=col, width=8, height=8,
                         corner_radius=4).pack(side="left", padx=2)
        ctk.CTkLabel(log_hdr, text="console",
                     font=("Segoe UI", 8), text_color=C["muted"],
                     ).pack(side="left", padx=4)

        self.log_text = scrolledtext.ScrolledText(
            log_card,
            bg=C["log_bg"], fg=C["log_fg"],
            font=FONT_MONO if self._font_exists("Cascadia Code") else FONT_LOG,
            insertbackground=C["log_fg"],
            wrap=tk.WORD, state=tk.DISABLED,
            relief=tk.FLAT, bd=0, padx=14, pady=10,
            selectbackground="#2D3748", height=6,
        )
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("ok",    foreground=C["log_ok"])
        self.log_text.tag_config("err",   foreground=C["log_err"])
        self.log_text.tag_config("warn",  foreground=C["log_warn"])
        self.log_text.tag_config("info",  foreground=C["log_info"])
        self.log_text.tag_config("muted", foreground=C["log_muted"])

        bottom_right = ctk.CTkFrame(right, fg_color="transparent")
        bottom_right.pack(fill="x", pady=(8, 0))

        self._section_title(bottom_right, "Validade do Laudo (dias)")
        card_val = self._card(bottom_right)
        val_row = ctk.CTkFrame(card_val, fg_color="transparent")
        val_row.pack(fill="x")
        self.entry_validade = ctk.CTkEntry(
            val_row, font=("Segoe UI", 13, "bold"),
            fg_color=C["surface2"], border_color=C["border"],
            text_color=C["accent"], border_width=1, corner_radius=8, height=44,
            placeholder_text="Ex: 12", placeholder_text_color=C["muted"],
            justify="center",
        )
        self.entry_validade.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self.entry_validade.bind("<Return>",     lambda e: self._salvar_validade())
        self.entry_validade.bind("<FocusOut>",   lambda e: self._salvar_validade())
        self.entry_validade.bind("<KeyRelease>", lambda e: self._salvar_validade())
        self.lbl_val_status = ctk.CTkLabel(val_row, text="",
                                            font=("Segoe UI", 10),
                                            text_color=C["ok"], width=80)
        self.lbl_val_status.pack(side="left")

        self._section_title(bottom_right, "Configurações")
        self._btn_salvar_cfg = ctk.CTkButton(
            bottom_right, text="💾  Salvar Config",
            font=FONT_BTN, fg_color=C["neut_bg"], hover_color=C["blue_lite"],
            text_color=C["text"], border_width=1, border_color=C["neut_brd"],
            corner_radius=10, height=42,
            command=self._salvar_config_manual,
        )
        self._btn_salvar_cfg.pack(fill="x")

        self._dot_phase = 0.0
        self._animate_dot()

    def _section_title(self, parent, text, top_pad=6):
        ctk.CTkLabel(parent, text=text,
                     font=("Segoe UI", 11, "bold"),
                     text_color=C["muted"], anchor="w",
                     ).pack(anchor="w", pady=(top_pad, 3))

    def _card(self, parent):
        card = ctk.CTkFrame(parent, fg_color=C["surface"],
                             border_color=C["border"], border_width=1,
                             corner_radius=8)
        card.pack(fill="x", pady=(0, 4))
        inner = ctk.CTkFrame(card, fg_color="transparent")
        inner.pack(fill="x", padx=14, pady=10)
        return inner

    def _font_exists(self, name: str) -> bool:
        try:
            import tkinter.font as tkfont
            return name in tkfont.families()
        except Exception:
            return False

    def _animate_dot(self):
        if self.executando:
            self._dot_phase = (self._dot_phase + 0.12) % (2 * math.pi)
            brightness = int(180 + 75 * math.sin(self._dot_phase))
            r = min(brightness, 255)
            g = min(int(brightness * 0.78), 255)
            self._dot_canvas.itemconfig(self._dot, fill=f"#{r:02x}{g:02x}00")
        else:
            self._dot_canvas.itemconfig(self._dot, fill=C["muted"])
        self.root.after(16, self._animate_dot)

    def _log_ui(self, msg: str):
        def _append():
            if msg.startswith("PROGRESSO:"):
                partes = msg.replace("PROGRESSO:", "").strip().split("/")
                if len(partes) == 2:
                    self.progress_var.set(
                        f"Processando {partes[0].strip()} de {partes[1].strip()}..."
                    )
                return

            if msg.startswith("__UUID__:"):
                self._ultimo_uuid = msg.replace("__UUID__:", "").strip()
                if self._validade_pendente and self.api:
                    self._enviar_validade(self._ultimo_uuid, self._validade_pendente)
                return

            if "✓" in msg:       tag = "ok"
            elif "✗" in msg or "ERRO" in msg or "ERROR" in msg: tag = "err"
            elif "AVISO" in msg or "CANCELAD" in msg:           tag = "warn"
            elif msg.startswith("=") or "INICIANDO" in msg or "FINALIZADO" in msg: tag = "info"
            elif not msg.strip(): tag = "muted"
            else:                 tag = None

            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n", tag or "")
            self.log_text.see(tk.END)
            self.log_text.configure(state=tk.DISABLED)

        self.root.after(0, _append)

    def _selecionar_pasta(self):
        pasta = filedialog.askdirectory(title="Selecione a pasta Comparativos")
        if pasta:
            self.entry_pasta.delete(0, tk.END)
            self.entry_pasta.insert(0, pasta)
            self._atualizar_subpastas(pasta)

    def _selecionar_credentials(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione credentials.json",
            filetypes=[("JSON", "*.json")],
        )
        if arquivo:
            self.entry_creds.delete(0, tk.END)
            self.entry_creds.insert(0, arquivo)

    def _atualizar_subpastas(self, pasta: str):
        try:
            p = Path(pasta)
            subs = []
            for d in sorted(p.iterdir()):
                if d.is_dir():
                    imgs = [f for f in d.iterdir() if f.suffix.lower() in IMAGE_EXTENSIONS]
                    if imgs:
                        subs.append(f"{d.name} ({len(imgs)} img)")
            if subs:
                self.lbl_subpastas.configure(
                    text="✓  " + "   ·   ".join(subs), text_color=C["ok"])
            else:
                self.lbl_subpastas.configure(
                    text="Nenhuma subpasta com imagens encontrada", text_color="#991B1B")
        except Exception:
            pass

    def _salvar_validade(self):
        valor = self.entry_validade.get().strip()
        if not valor:
            return
        if not valor.isdigit():
            self.lbl_val_status.configure(text="✗ inválido", text_color=C["err"])
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))
            return
        self._validade_pendente = valor
        if self._ultimo_uuid and self.api:
            self._enviar_validade(self._ultimo_uuid, valor)
        else:
            self.lbl_val_status.configure(text="💾 salvo", text_color=C["muted"])
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))

    def _enviar_validade(self, uuid: str, valor: str):
        def _run():
            try:
                res = self.api.editar_avaliacao(uuid, VALIDADELAUDO=valor)
                cor = C["ok"] if res.get("status") == "sucesso" else C["err"]
                txt = "✓ salvo" if res.get("status") == "sucesso" else "✗ erro"
                self.root.after(0, lambda: self.lbl_val_status.configure(text=txt, text_color=cor))
            except Exception:
                self.root.after(0, lambda: self.lbl_val_status.configure(text="✗ erro", text_color=C["err"]))
            self.root.after(2000, lambda: self.lbl_val_status.configure(text=""))
        Thread(target=_run, daemon=True).start()

    def _executar(self):
        if self.executando:
            return

        pasta = self.entry_pasta.get().strip()
        if not pasta or not os.path.isdir(pasta):
            messagebox.showerror("Erro", "Selecione uma pasta válida.")
            return

        creds = self.entry_creds.get().strip()
        if not creds or not os.path.isfile(creds):
            messagebox.showerror("Erro", "Selecione o arquivo credentials.json.")
            return

        planilha_id = self.entry_planilha.get().strip()
        if not planilha_id:
            messagebox.showerror("Erro", "Informe o ID da planilha Google Sheets.")
            return

        self.config.sistema            = self.var_sistema.get()
        self.config.pasta_comparativos = pasta
        self.config.credentials_path   = creds
        self.config.planilha_id        = planilha_id
        self.config.excluir_imagens    = self.var_excluir.get()

        self.executando = True
        self.btn_executar.configure(state="disabled", fg_color=C["accent_hov"])
        self.btn_cancelar.configure(state="normal", fg_color="#8B1A1A",
                                     text_color="#FFAAAA", hover_color="#B91C1C")
        self.progress_var.set("Iniciando...")
        self._salvar_config()

        def _run():
            try:
                self.robo = RoboFechamento(
                    config=self.config,
                    logger=self.logger,
                    callback_progresso=self._log_ui,
                    callback_confirmar=None,
                    callback_validade=lambda: self.entry_validade.get().strip(),
                )
                sistema_cfg = SISTEMAS[self.config.sistema]
                self.api = MultiAAPI(sistema_cfg, self.logger)
                self.robo.api = self.api
                self.robo.executar()
            finally:
                self.root.after(0, self._finalizar_execucao)

        Thread(target=_run, daemon=True).start()

    def _cancelar(self):
        if self.robo:
            self.robo.cancelar()
            self._log_ui("Cancelamento solicitado...")

    def _finalizar_execucao(self):
        self.executando = False
        self.btn_executar.configure(state="normal", fg_color=C["accent"])
        self.btn_cancelar.configure(state="disabled", fg_color=C["neut_bg"],
                                     text_color=C["muted"])
        self.progress_var.set("✓  Finalizado")

    def _salvar_config(self):
        path = _base_dir() / "config.json"
        data = {
            "sistema":          self.config.sistema,
            "planilha_id":      self.config.planilha_id,
            "credentials_path": self.config.credentials_path,
            "pasta":            self.config.pasta_comparativos,
            "excluir_imagens":  self.config.excluir_imagens,
            "validade_laudo":   self.entry_validade.get().strip(),
        }
        with open(path, "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def _salvar_config_manual(self):
        self.config.sistema            = self.var_sistema.get()
        self.config.planilha_id        = self.entry_planilha.get().strip()
        self.config.credentials_path   = self.entry_creds.get().strip()
        self.config.pasta_comparativos = self.entry_pasta.get().strip()
        self.config.excluir_imagens    = self.var_excluir.get()
        self._salvar_config()
        self._btn_salvar_cfg.configure(text="✓  Config Salva!",
                                        fg_color=C["ok"], text_color="#000000",
                                        state="disabled")
        self.root.after(1500, lambda: self._btn_salvar_cfg.configure(
            text="💾  Salvar Config", fg_color=C["neut_bg"],
            text_color=C["text"], state="normal",
        ))

    def _carregar_config(self):
        path = _base_dir() / "config.json"
        if not path.exists():
            return
        try:
            with open(path, encoding="utf-8") as f:
                data = json.load(f)

            self.var_sistema.set(data.get("sistema", "MultiA Mais"))

            if data.get("planilha_id"):
                self.entry_planilha.delete(0, tk.END)
                self.entry_planilha.insert(0, data["planilha_id"])

            if data.get("credentials_path"):
                self.entry_creds.delete(0, tk.END)
                self.entry_creds.insert(0, data["credentials_path"])

            pasta = data.get("pasta", "")
            if pasta and os.path.isdir(pasta):
                self.entry_pasta.delete(0, tk.END)
                self.entry_pasta.insert(0, pasta)
                self._atualizar_subpastas(pasta)

            self.var_excluir.set(data.get("excluir_imagens", True))

            validade = data.get("validade_laudo", "")
            if validade:
                self.entry_validade.delete(0, tk.END)
                self.entry_validade.insert(0, validade)

        except Exception as e:
            self.logger.warning(f"Erro ao carregar config: {e}")

    def run(self):
        self.root.mainloop()


# ============================================================
# MAIN
# ============================================================

if __name__ == "__main__":
    app = App()
    app.run()
