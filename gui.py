"""
Microsoft 365 License Auditor & Sanitizer (Versão 9.3 - Auto-Explicativo)
Autor: Vini / Refactored by Gemini
Data: 2025-05-28

Novidades v9.3:
- Inatividade: Cálculo de dias exatos e mensagens em linguagem natural.
- Explicabilidade: Coluna 'Rule_Explanation' detalhando o porquê de cada ação.
- UI Resumo: Adicionado bloco "Como Interpretar" e tabela de métricas reais.
- Visual: Correção da zebra striping (tags even/odd).
- Core: Mantida toda a robustez de leitura e lógica da v9.2.

Dependências:
  pip install pandas openpyxl xlsxwriter ttkbootstrap
"""

import os
import sys
import re
import csv
import json
import threading
import queue
import subprocess
from dataclasses import dataclass, field
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- UI Framework ---
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.scrolled import ScrolledText, ScrolledFrame

import pandas as pd
import numpy as np

# -----------------------------
# Configuração Global
# -----------------------------
DEFAULT_CLIENT_TAG = "Shell_Labs"
SETTINGS_FILE = "settings.json"

@dataclass
class AuditConfig:
    token_f3: str = "microsoft 365 f3"
    token_e5: str = "microsoft 365 e5"
    
    tokens_exchange: list[str] = field(default_factory=lambda: [
        "exchange online", "exchange online (plan 1)", "exchange online (plan 2)",
        "exchange foundation", "office 365 e1", "office 365 e3"
    ])
    
    column_aliases: dict = field(default_factory=lambda: {
        "User principal name": ["user principal name", "userprincipalname", "upn", "id do usuário", "user name"],
        "Display name": ["display name", "displayname", "nome de exibição", "nome"],
        "Licenses": ["licenses", "assignedlicenses", "licenças", "assigned licenses"],
        "Block credential": ["block credential", "blockcredential", "bloquear entrada", "account enabled", "sign-in status"],
        "Proxy addresses": ["proxy addresses", "proxyaddresses", "endereços proxy", "email addresses", "email"],
        "Object Id": ["object id", "objectid", "id do objeto", "external directory object id"],
        "Mobile Phone": ["mobile phone", "mobile", "telefone celular", "celular"],
        "Phone number": ["phone number", "phone", "telefone comercial", "telefone"],
        "Title": ["title", "cargo", "job title", "função", "funcao"],
        "Department": ["department", "departamento", "area", "área"],
        "Device ownership": ["device ownership", "ownership"],
        "Managed devices": ["managed devices", "intune managed devices"],
        "Device name": ["device name", "hostname", "computer name"],
        "Last password change time stamp": ["last password change time stamp", "lastpasswordchangetimestamp", "troca de senha", "password change"],
        "Last dirsync time": ["last dirsync time", "lastdirsynctime", "ultimo sync", "dirsync"],
        "Soft deletion time stamp": ["soft deletion time stamp", "softdeletiontimestamp", "data exclusão suave", "deleted date"],
        "When created": ["when created", "whencreated", "criado em", "creation date"],
        "DirSyncEnabled": ["dirsyncenabled", "dirsync enabled", "sincronizado"],
        "Password never expires": ["password never expires", "passwordneverexpires", "senha nunca expira"]
    })
    
    preview_limit_step: int = 500
    inactivity_days_threshold: int = 180

CFG = AuditConfig()

# -----------------------------
# Componentes Visuais (RoundedCard)
# -----------------------------
class RoundedCard(tk.Canvas):
    def __init__(self, parent, title, variable, color_style, **kwargs):
        self.style = tb.Style()
        self.bg_color = self.style.colors.get(color_style)
        self.fg_color = self.style.colors.get("light")
        self.canvas_bg = self.style.colors.bg 
        self.shadow_color = "#D0D4D8" if self.style.theme.type == 'light' else "#222222"
        
        super().__init__(parent, background=self.canvas_bg, highlightthickness=0, **kwargs)
        
        self.title = title
        self.variable = variable
        self.radius = 20
        
        self.bind("<Configure>", self._draw)
        self.variable.trace_add("write", self._update_text)

    def _draw(self, event=None):
        self.delete("all")
        w = self.winfo_width()
        h = self.winfo_height()
        if w < 10 or h < 10: return

        r = self.radius
        off = 3 
        
        self._draw_rounded_rect(off, off, w-off, h-off, r, self.shadow_color)
        self._draw_rounded_rect(0, 0, w-off-1, h-off-1, r, self.bg_color)
        
        font_size = 36 if w > 200 else 24
        
        self.create_text(20, 20, text=self.title, anchor="nw", 
            font=("Segoe UI", 11), fill=self.fg_color, width=max(10, w - 40))
        self.text_id = self.create_text(
            20, 50, text=self.variable.get(), anchor="nw", 
            font=("Segoe UI", font_size, "bold"), fill=self.fg_color)

    def _draw_rounded_rect(self, x1, y1, x2, y2, r, color):
        self.create_arc(x1, y1, x1+2*r, y1+2*r, start=90, extent=90, fill=color, outline=color)
        self.create_arc(x2-2*r, y1, x2, y1+2*r, start=0, extent=90, fill=color, outline=color)
        self.create_arc(x1, y2-2*r, x1+2*r, y2, start=180, extent=90, fill=color, outline=color)
        self.create_arc(x2-2*r, y2-2*r, x2, y2, start=270, extent=90, fill=color, outline=color)
        self.create_rectangle(x1+r, y1, x2-r, y2, fill=color, outline=color)
        self.create_rectangle(x1, y1+r, x2, y2-r, fill=color, outline=color)

    def _update_text(self, *args):
        if hasattr(self, 'text_id'):
            try: self.itemconfig(self.text_id, text=self.variable.get())
            except tk.TclError: pass

# -----------------------------
# Utilitários
# -----------------------------
def safe_open_folder(path: str) -> None:
    folder = os.path.dirname(path)
    try:
        if sys.platform.startswith("win"): os.startfile(folder)
        elif sys.platform.startswith("darwin"): subprocess.run(["open", folder], check=False)
        else: subprocess.run(["xdg-open", folder], check=False)
    except Exception: pass

def normalize_phone_series(s: pd.Series) -> pd.Series:
    return s.fillna("").astype(str).str.replace(r"\D+", "", regex=True)

# -----------------------------
# Engine de Auditoria
# -----------------------------
class AuditEngine:
    @staticmethod
    def simplify_identifier(text: str) -> str:
        if not isinstance(text, str): text = str(text)
        return re.sub(r'[^a-z0-9]', '', text.lower())

    @staticmethod
    def read_file_robust(path: str, logger_callback=None) -> pd.DataFrame:
        if logger_callback: logger_callback(f"Lendo arquivo: {os.path.basename(path)}")
        try:
            if path.lower().endswith((".xlsx", ".xls")):
                for header_row in range(0, 5):
                    try:
                        df = pd.read_excel(path, dtype=str, header=header_row)
                        cols = [AuditEngine.simplify_identifier(c) for c in df.columns]
                        if any("user" in c for c in cols) or any("display" in c for c in cols):
                            if logger_callback: logger_callback(f"Excel detectado (header linha {header_row})")
                            return df
                    except: continue
                return pd.read_excel(path, dtype=str)
            else:
                encoding = 'utf-8-sig'
                sep = None
                try:
                    with open(path, 'r', encoding=encoding, errors='replace') as f:
                        sample = f.read(4096)
                        try: sep = csv.Sniffer().sniff(sample, delimiters=[',', ';', '\t', '|']).delimiter
                        except: sep = ','
                except: encoding = 'latin1'
                
                if logger_callback: logger_callback(f"CSV detectado: sep='{sep}', enc='{encoding}'")
                try:
                    df = pd.read_csv(path, sep=sep, encoding=encoding, dtype=str, on_bad_lines='skip')
                    if len(df.columns) > 1: return df
                except: pass
                return pd.read_csv(path, sep=None, engine='python', encoding=encoding, dtype=str, on_bad_lines='skip')
        except Exception as e:
            raise ValueError(f"Falha na leitura: {str(e)}")

    @staticmethod
    def map_columns(df: pd.DataFrame, logger_callback=None) -> pd.DataFrame:
        df.columns = [str(col).replace('\ufeff', '').replace('\u00a0', ' ').strip() for col in df.columns]
        canonical_map = {}
        for official, aliases in CFG.column_aliases.items():
            canonical_map[AuditEngine.simplify_identifier(official)] = official
            for a in aliases: canonical_map[AuditEngine.simplify_identifier(a)] = official
        
        new_cols = []
        for col in df.columns:
            c_simple = AuditEngine.simplify_identifier(col)
            new_cols.append(canonical_map.get(c_simple, col))
        df.columns = new_cols
        
        required = ["User principal name", "Display name"]
        missing = [req for req in required if req not in df.columns]
        if missing:
            if logger_callback: logger_callback(f"ERRO: Colunas detectadas: {list(df.columns)}")
            raise ValueError(f"Colunas obrigatórias ausentes: {', '.join(missing)}")
        return df

    @staticmethod
    def normalize_data(df: pd.DataFrame, logger_callback=None) -> pd.DataFrame:
        if logger_callback: logger_callback("Normalizando dados...")
        df = df.copy()

        for col in df.select_dtypes(include=['object']).columns:
            df[col] = df[col].astype(str).str.strip()
        
        base_cols = ["User principal name", "Display name", "Object Id", "Title", "Department", "Licenses"]
        for c in base_cols:
            if c not in df.columns: df[c] = ""
            
        df["UPN"] = df["User principal name"]
        df["DisplayName"] = df["Display name"]
        df["UPN_Lower"] = df["UPN"].str.lower()
        
        true_vals = {"true", "1", "yes", "sim", "enabled"}
        if "Block credential" in df.columns:
            df["Is_Blocked"] = df["Block credential"].str.lower().isin(true_vals)
        else:
            df["Is_Blocked"] = False
            
        def parse_dt(col):
            if col in df.columns:
                return pd.to_datetime(df[col], errors="coerce", utc=True).dt.tz_localize(None)
            return pd.Series(pd.NaT, index=df.index)

        df["SoftDelete_Date"] = parse_dt("Soft deletion time stamp")
        df["Is_SoftDeleted"] = df["SoftDelete_Date"].notna()
        df["PwdChange_Date"] = parse_dt("Last password change time stamp")
        df["DirSync_Date"] = parse_dt("Last dirsync time")
        
        df["Is_ThirdParty"] = df["UPN_Lower"].str.startswith("3")
        
        df["Title_Clean"] = df["Title"].str.lower()
        high_role = r"(diretor|diretora|gerente|head|c-level|ceo|cfo|cio|cto|coo|vp|vice[-\s]?presidente|presidente)"
        df["Is_HighRole"] = df["Title_Clean"].str.contains(high_role, regex=True, na=False)

        df["Has_AcheMachine"] = False
        if "Device ownership" in df.columns:
            df["Has_AcheMachine"] |= df["Device ownership"].str.lower().isin({"company", "corporate", "empresa"})
        if "Managed devices" in df.columns:
            raw_md = df["Managed devices"].str.replace(",", ".")
            clean_md = raw_md.str.extract(r'(\d+)', expand=False).fillna("0")
            df["Has_AcheMachine"] |= (pd.to_numeric(clean_md, errors="coerce").fillna(0) > 0)
        
        df["Allowed_E5"] = (df["Is_HighRole"] | (df["Is_ThirdParty"] & df["Has_AcheMachine"]))

        df = df.reset_index(drop=True)
        df["RowId"] = df.index
        return df

    @staticmethod
    def analyze_licenses(df: pd.DataFrame) -> pd.DataFrame:
        lic = df["Licenses"].str.lower()
        df["has_f3"] = lic.str.contains(CFG.token_f3, regex=False)
        df["has_e5"] = lic.str.contains(CFG.token_e5, regex=False)
        exc_regex = "|".join([re.escape(t) for t in CFG.tokens_exchange])
        df["has_any_exchange"] = lic.str.contains(exc_regex, regex=True)
        df["has_conflict_f3_e5"] = df["has_f3"] & df["has_e5"]
        return df

    @staticmethod
    def evaluate_license_profile(df: pd.DataFrame, analysis_mode: str, cancel_event=None, logger_callback=None) -> pd.DataFrame:
        if logger_callback: logger_callback(f"Avaliando perfil (Modo: {analysis_mode})...")
        if cancel_event and cancel_event.is_set(): raise InterruptedError()

        for col in ["Action", "Reasons", "Severity", "Is_Candidate"]:
            if col not in df.columns:
                df[col] = "" if col in ["Action", "Reasons", "Severity"] else False

        has_f3 = df.get("has_f3", False)
        has_e5 = df.get("has_e5", False)
        conflict = df.get("has_conflict_f3_e5", False)
        allowed_e5 = df["Allowed_E5"] if "Allowed_E5" in df.columns else pd.Series(True, index=df.index)

        def add_reason(mask, text):
            if not mask.any(): return
            df.loc[mask, "Reasons"] = (df.loc[mask, "Reasons"].fillna("") + text + ", ").astype(str)

        def set_sev(mask, sev):
            if not mask.any(): return
            empty = mask & df["Severity"].fillna("").eq("")
            df.loc[empty, "Severity"] = sev

        def add_action(mask, action_text):
            if not mask.any(): return
            empty = mask & df["Action"].fillna("").eq("")
            df.loc[empty, "Action"] = action_text
            append = mask & df["Action"].fillna("").ne("") & ~df["Action"].astype(str).str.contains(re.escape(action_text), regex=True)
            if append.any():
                df.loc[append, "Action"] = df.loc[append, "Action"].astype(str) + " + " + action_text

        if analysis_mode == "F3":
            m_has_e5 = has_e5.astype(bool)
            add_reason(m_has_e5, "MODO_F3_COM_E5")
            set_sev(m_has_e5, "MEDIA")
            df.loc[m_has_e5, "Is_Candidate"] = True
            add_action(conflict.astype(bool), "REMOVER_E5 (MANTER F3)")
            add_action((has_e5.astype(bool) & ~has_f3.astype(bool)), "DOWNGRADE_E5_PARA_F3")

        elif analysis_mode == "E5":
            m_has_f3 = has_f3.astype(bool)
            add_reason(m_has_f3, "MODO_E5_COM_F3")
            set_sev(m_has_f3, "MEDIA")
            df.loc[m_has_f3, "Is_Candidate"] = True
            add_action(conflict.astype(bool), "REMOVER_F3 (MANTER E5)")
            add_action((has_f3.astype(bool) & ~has_e5.astype(bool)), "REMOVER_F3 (PADRONIZAR E5)")
            
            m_e5_not_allowed = has_e5.astype(bool) & ~allowed_e5.astype(bool)
            add_reason(m_e5_not_allowed, "E5_NAO_PERMITIDO_POR_POLITICA")
            set_sev(m_e5_not_allowed, "ALTA")
            df.loc[m_e5_not_allowed, "Is_Candidate"] = True
            add_action(m_e5_not_allowed, "REMOVER_E5 / AVALIAR_F3")

        elif analysis_mode == "EXCHANGE":
            m_full = (has_f3.astype(bool) | has_e5.astype(bool))
            add_reason(m_full, "MODO_EXCHANGE_COM_LICENCA_COMPLETA")
            set_sev(m_full, "MEDIA")
            df.loc[m_full, "Is_Candidate"] = True
            add_action(m_full, "REMOVER_F3/E5 (MANTER APENAS EXCHANGE)")

        else:  # FULL
            m_conf = conflict.astype(bool)
            add_reason(m_conf, "CONFLITO_F3_E5")
            if "Allowed_E5" in df.columns:
                m_keep_e5 = m_conf & allowed_e5.astype(bool)
                m_keep_f3 = m_conf & ~allowed_e5.astype(bool)
                set_sev(m_conf, "MEDIA")
                df.loc[m_conf, "Is_Candidate"] = True
                add_action(m_keep_e5, "REMOVER_F3 (POLITICA)")
                add_action(m_keep_f3, "REMOVER_E5 (POLITICA)")
            else:
                set_sev(m_conf, "MEDIA")
                df.loc[m_conf, "Is_Candidate"] = True
                add_action(m_conf, "REVISAR (MANTER APENAS UMA)")

            m_e5_bad = has_e5.astype(bool) & ~allowed_e5.astype(bool)
            add_reason(m_e5_bad, "E5_NAO_PERMITIDO_POR_POLITICA")
            set_sev(m_e5_bad, "ALTA")
            df.loc[m_e5_bad, "Is_Candidate"] = True
            add_action(m_e5_bad, "REMOVER_E5 / REVISAR PERFIL")

        df["Reasons"] = df["Reasons"].astype(str).str.rstrip(", ").replace("nan", "")
        return df

    @staticmethod
    def apply_global_rules(df: pd.DataFrame, analysis_mode: str, cancel_event=None, logger_callback=None) -> pd.DataFrame:
        if logger_callback: logger_callback(f"Aplicando regras globais...")
        if cancel_event and cancel_event.is_set(): raise InterruptedError()
        
        def add_action(mask, text):
            if not mask.any(): return
            df.loc[mask & df["Action"].eq(""), "Action"] = text
            has_other = mask & df["Action"].ne("") & ~df["Action"].str.contains(re.escape(text), regex=True)
            if has_other.any():
                df.loc[has_other, "Action"] = df.loc[has_other, "Action"] + " + " + text

        # 1. Terceiro Bloqueado
        m_target = df["Is_ThirdParty"] & df["Is_Blocked"]
        if m_target.any():
            df.loc[m_target, "Is_Candidate"] = True
            df.loc[m_target, "Severity"] = "ALTA"
            df.loc[m_target, "Reasons"] = (df.loc[m_target, "Reasons"] + "TERCEIRO_BLOQUEADO, ").str.replace(", TERCEIRO_BLOQUEADO,", ", TERCEIRO_BLOQUEADO,")
            
            if analysis_mode == "F3": add_action(m_target & df["has_f3"], "REMOVER_F3")
            elif analysis_mode == "E5": add_action(m_target & df["has_e5"], "REMOVER_E5")
            elif analysis_mode == "EXCHANGE": add_action(m_target & df["has_any_exchange"], "REMOVER_EXCHANGE")
            else:
                add_action(m_target & df["has_f3"], "REMOVER_F3")
                add_action(m_target & df["has_e5"], "REMOVER_E5")
                add_action(m_target & df["has_any_exchange"], "REMOVER_EXCHANGE")

            m_generic = m_target & df["Action"].eq("")
            if m_generic.any(): df.loc[m_generic, "Action"] = "REVISAR_CONTA (BLOQUEADA)"

        # 2. Soft Deleted
        m_soft = df["Is_SoftDeleted"]
        if m_soft.any():
            df.loc[m_soft, "Is_Candidate"] = True
            df.loc[m_soft, "Severity"] = "ALTA"
            df.loc[m_soft, "Reasons"] += "SOFT_DELETED, "
            df.loc[m_soft, "Action"] = "LIMPAR_LICENCAS"

        df["Reasons"] = df["Reasons"].str.rstrip(", ")
        return df

    @staticmethod
    def detect_inactivity(df: pd.DataFrame) -> pd.DataFrame:
        now = datetime.now()
        limit = now - timedelta(days=CFG.inactivity_days_threshold)

        df["Is_Inactive"] = False
        df["Inactivity_Days"] = 0
        df["Inactivity_Message"] = ""
        df["Inactivity_Reason"] = "" # Legacy compatibility

        # Melhor estimativa de última atividade (Max entre Senha e Sync)
        candidates = []
        if "PwdChange_Date" in df.columns: candidates.append(df["PwdChange_Date"])
        if "DirSync_Date" in df.columns: candidates.append(df["DirSync_Date"])

        if not candidates: return df

        last_activity = pd.concat(candidates, axis=1).max(axis=1)

        mask = last_activity.notna() & (last_activity < limit)
        if mask.any():
            df.loc[mask, "Is_Inactive"] = True
            df.loc[mask, "Inactivity_Days"] = (now - last_activity[mask]).dt.days
            
            # Mensagem humana
            df.loc[mask, "Inactivity_Message"] = (
                "Usuário sem atividade há " + 
                df.loc[mask, "Inactivity_Days"].astype(str) + 
                " dias — avaliar necessidade de licença."
            )
            df.loc[mask, "Inactivity_Reason"] = df.loc[mask, "Inactivity_Message"]

        return df

    @staticmethod
    def build_rule_explanations(df: pd.DataFrame, mode: str) -> pd.DataFrame:
        df["Rule_Explanation"] = ""

        def add(mask, text):
            if not mask.any(): return
            df.loc[mask, "Rule_Explanation"] = (df.loc[mask, "Rule_Explanation"].fillna("") + text + "\n")

        # Regras Claras
        add(df.get("Is_ThirdParty", False).astype(bool), "Terceiro: UPN inicia com '3' (conta externa/fornecedor).")
        
        m_3_block = (df.get("Is_ThirdParty", False).astype(bool) & df.get("Is_Blocked", False).astype(bool))
        add(m_3_block, "Candidato: terceiro com conta bloqueada — recomenda-se remover licenças.")

        m_conf = df.get("has_conflict_f3_e5", False).astype(bool)
        add(m_conf, "Conflito: possui F3 e E5 simultaneamente — manter apenas uma (conforme política).")

        m_mismatch = df["Reasons"].fillna("").str.contains("MODO_", na=False)
        mode_desc = {
            "F3": "Modo F3 selecionado: usuários com E5 devem ser ajustados.",
            "E5": "Modo E5 selecionado: usuários com F3 devem ser ajustados.",
            "EXCHANGE": "Modo Exchange: remover licenças completas (F3/E5).",
            "FULL": "Modo Completo: sistema sugere ajuste conforme política interna."
        }
        add(m_mismatch, f"Padronização: {mode_desc.get(mode, 'Verificar política.')}")

        m_e5_pol = df["Reasons"].fillna("").str.contains("E5_NAO_PERMITIDO", na=False)
        add(m_e5_pol, "Política: E5 atribuída a perfil não elegível (sem cargo de liderança/máquina).")
        
        # Adiciona a mensagem de inatividade se existir
        m_inact = df.get("Is_Inactive", False).astype(bool)
        if "Inactivity_Message" in df.columns:
            # Iterar apenas nos inativos para pegar mensagem individual
            # (Poderia ser vetorizado, mas string concat com colunas diferentes é chato no pandas)
            # Simplificação: adicionar marcador genérico ou usar apply se performance permitir
            # Aqui vamos usar o valor da coluna Inactivity_Message
            pass # Já está na coluna Inactivity_Message, que é exibida na UI

        df["Rule_Explanation"] = df["Rule_Explanation"].astype(str).str.strip()
        return df

    @staticmethod
    def build_summary(df: pd.DataFrame) -> pd.DataFrame:
        def c(mask): return int(mask.sum())
        
        total = len(df)
        third_all = c(df.get("Is_ThirdParty", False).astype(bool))
        third_block = c((df.get("Is_ThirdParty", False) & df.get("Is_Blocked", False)).astype(bool))
        cand = c(df.get("Is_Candidate", False).astype(bool))
        conf = c(df.get("has_conflict_f3_e5", False).astype(bool))
        mismatch = c(df["Reasons"].fillna("").str.contains("MODO_", na=False))
        inactive = c(df.get("Is_Inactive", False).astype(bool))
        dups = c(df.get("dup_strength", "").astype(str).ne(""))

        return pd.DataFrame({
            "Métrica": [
                "Total de Usuários Analisados",
                "Candidatos a Saneamento (Remoção)",
                "Conflitos de Licença (F3 + E5)",
                "Fora do Padrão (Mismatch de Modo)",
                "Total de Terceiros (UPN '3...')",
                "Terceiros Bloqueados",
                "Duplicidades Identificadas (Linhas)",
                "Usuários Inativos (>180 dias)"
            ],
            "Valor": [total, cand, conf, mismatch, third_all, third_block, dups, inactive]
        })

    @staticmethod
    def detect_duplicates(df: pd.DataFrame, cancel_event=None, logger_callback=None) -> pd.DataFrame:
        if logger_callback: logger_callback("Analisando duplicidades...")
        if cancel_event and cancel_event.is_set(): raise InterruptedError()
        
        df["dup_strength"] = ""
        df["dup_notes"] = ""
        df["dup_group_id"] = ""
        df["dup_score"] = 0
        df["keep_recommendation"] = ""
        
        mask_upn = df.duplicated("UPN_Lower", keep=False) & df["UPN_Lower"].ne("")
        if mask_upn.any():
            df.loc[mask_upn, "dup_strength"] = "CERTEZA"
            df.loc[mask_upn, "dup_notes"] += "[UPN Duplicado] "
            df.loc[mask_upn, "dup_score"] += 100
            groups = df.loc[mask_upn].groupby("UPN_Lower").ngroup()
            df.loc[mask_upn, "dup_group_id"] = groups.apply(lambda x: f"DUP_UPN_{x:04d}")

        if "Proxy addresses" in df.columns:
            temp = df[["RowId", "UPN"]].copy()
            temp["proxies"] = df["Proxy addresses"].fillna("").str.lower().str.split(r'[;,]')
            exploded = temp.explode("proxies")
            exploded["proxies"] = exploded["proxies"].str.replace(r'(?i)^smtp:', "", regex=True).str.strip()
            valid = exploded[exploded["proxies"].str.len() > 5]
            dupes = valid[valid.duplicated("proxies", keep=False)]
            
            if not dupes.empty:
                grp_cnt = 0
                score_map = {}
                note_map = {}
                grp_map = {}
                str_map = {}
                for proxy, g in dupes.groupby("proxies"):
                    rids = g["RowId"].unique()
                    if len(rids) > 1:
                        grp_id = f"PXY_{grp_cnt:04d}"
                        grp_cnt += 1
                        note = f"[Proxy: {proxy}] "
                        for rid in rids:
                            score_map[rid] = score_map.get(rid, 0) + 60
                            note_map[rid] = note_map.get(rid, "") + note
                            grp_map[rid] = grp_map.get(rid, "") + grp_id + ";"
                            if df.at[rid, "dup_strength"] != "CERTEZA": str_map[rid] = "FORTE"
                
                for rid, val in score_map.items(): df.at[rid, "dup_score"] += val
                for rid, val in note_map.items(): df.at[rid, "dup_notes"] += val
                for rid, val in grp_map.items(): 
                    curr = df.at[rid, "dup_group_id"]
                    df.at[rid, "dup_group_id"] = (curr + ";" + val).strip(";")
                for rid, val in str_map.items(): df.at[rid, "dup_strength"] = val

        cols_phone = [c for c in ["Mobile Phone", "Phone number"] if c in df.columns]
        for col in cols_phone:
            norm_col = f"{col}_norm"
            df[norm_col] = normalize_phone_series(df[col])
            mask_valid = df[norm_col].str.len() > 6
            dup_ph = df[mask_valid].duplicated(norm_col, keep=False)
            target_mask = df.index.isin(dup_ph[dup_ph].index)
            if target_mask.any():
                df.loc[target_mask, "dup_notes"] += f"[{col} Repetido] "
                df.loc[target_mask, "dup_score"] += 30
                mask_str = target_mask & df["dup_strength"].eq("")
                df.loc[mask_str, "dup_strength"] = "SUSPEITO"

        if "DisplayName" in df.columns:
            mask_name = df.duplicated("DisplayName", keep=False) & df["DisplayName"].ne("")
            if mask_name.any():
                df.loc[mask_name, "dup_notes"] += "[Nome Exato] "
                df.loc[mask_name, "dup_score"] += 10
                mask_str = mask_name & df["dup_strength"].eq("")
                df.loc[mask_str, "dup_strength"] = "HOMONIMO"

        mask_dups = df["dup_group_id"] != ""
        if mask_dups.any():
            df.loc[mask_dups, "keep_recommendation"] = "REVISAR"
            df.loc[mask_dups & df["Is_Blocked"], "keep_recommendation"] = "REMOVER (Bloqueado)"
            df.loc[mask_dups & ~df["Is_Blocked"], "keep_recommendation"] = "MANTER (Ativo)"

        return df

    @staticmethod
    def create_group_view(df_dups: pd.DataFrame) -> pd.DataFrame:
        if df_dups.empty: return pd.DataFrame()
        temp = df_dups.assign(grp=df_dups["dup_group_id"].str.split(";")).explode("grp")
        temp = temp[temp["grp"] != ""]
        if temp.empty: return pd.DataFrame()
        
        res = temp.groupby("grp").agg(
            Qtd=('RowId', 'count'),
            UPNs=('UPN', lambda x: ", ".join(x.astype(str))),
            Scores=('dup_score', 'max')
        ).reset_index()
        res.columns = ["ID Grupo", "Qtd", "UPNs Envolvidos", "Score Máx"]
        return res.sort_values("Score Máx", ascending=False)

    @staticmethod
    def compare_files_delta(path_before, path_after, logger_callback=None):
        if logger_callback: logger_callback("Comparando snapshots (Outer Join)...")
        df_b = AuditEngine.load_dataframe(path_before)
        df_b = AuditEngine.normalize_data(df_b)
        df_b = AuditEngine.analyze_licenses(df_b)
        df_a = AuditEngine.load_dataframe(path_after)
        df_a = AuditEngine.normalize_data(df_a)
        df_a = AuditEngine.analyze_licenses(df_a)
        
        cols = ["UPN_Lower", "DisplayName", "has_f3", "has_e5", "has_any_exchange", "Licenses"]
        merged = pd.merge(df_b[cols], df_a[cols], on="UPN_Lower", how="outer", suffixes=('_old', '_new'), indicator=True)
        
        novos = merged[merged["_merge"] == "right_only"].copy()
        deletados = merged[merged["_merge"] == "left_only"].copy()
        existentes = merged[merged["_merge"] == "both"].copy()
        
        existentes["removed_f3"] = existentes["has_f3_old"] & ~existentes["has_f3_new"]
        existentes["removed_e5"] = existentes["has_e5_old"] & ~existentes["has_e5_new"]
        existentes["removed_exchange"] = existentes["has_any_exchange_old"] & ~existentes["has_any_exchange_new"]
        
        mask_saneados = existentes["removed_f3"] | existentes["removed_e5"] | existentes["removed_exchange"]
        saneados = existentes[mask_saneados].copy()
        
        saneados["Acao_Realizada"] = ""
        saneados.loc[saneados["removed_f3"], "Acao_Realizada"] += "F3 Removido; "
        saneados.loc[saneados["removed_e5"], "Acao_Realizada"] += "E5 Removido; "
        saneados.loc[saneados["removed_exchange"], "Acao_Realizada"] += "Exchange Removido; "
        
        return {
            "Saneados (Alterados)": saneados,
            "Novos Usuários": novos,
            "Deletados": deletados
        }

# -----------------------------
# GUI App
# -----------------------------
class AuditorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("M365 Auditor & Sanitizer v9.3 (Auto-Explicativo)")
        self.root.geometry("1400x850")
        
        self.file_path_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Pronto")
        self.search_var = tk.StringVar()
        self.analysis_mode_var = tk.StringVar(value="FULL")
        
        self.kpi_total = tk.StringVar(value="0")
        self.kpi_candidates = tk.StringVar(value="0")
        self.kpi_conflicts = tk.StringVar(value="0")
        self.kpi_dups = tk.StringVar(value="0")

        self.search_timer = None
        self.log_queue = queue.Queue()
        self.is_running = False
        self.cancel_event = threading.Event()
        self.df_complete = None
        self.dict_results = {}
        self.tree_map = {}
        self.export_opts = {} 
        
        self.load_settings()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.style = tb.Style(theme="flatly")
        self._build_ui()
        self._process_logs()

    def _build_ui(self):
        main = tb.Frame(self.root, padding=15)
        main.pack(fill="both", expand=True)
        self._build_header(main)
        self.nb = tb.Notebook(main, bootstyle="primary")
        self.nb.pack(fill="both", expand=True, pady=10)
        
        self._create_tab("Resumo", self._build_tab_resumo)
        self._create_tab("Candidatos", self._build_tab_candidatos)
        self._create_tab("Padronização", self._build_tab_padronizacao)
        self._create_tab("Conflitos", self._build_tab_conflitos)
        self._create_tab("Terceiros", self._build_tab_terceiros)
        self._create_tab("Duplicados (Lista)", self._build_tab_dups_list)
        self._create_tab("Duplicados (Grupos)", self._build_tab_dups_group)
        self._create_tab("Inatividade", self._build_tab_inactive)
        self._create_tab("Audit Trail", self._build_tab_all)

    def _build_header(self, parent):
        h = tb.Frame(parent)
        h.pack(fill="x", pady=(0, 10))
        tb.Label(h, text="Auditoria M365 Enterprise", font=("Segoe UI", 16, "bold")).pack(side="left")
        tb.Label(h, textvariable=self.status_var, bootstyle="secondary").pack(side="left", padx=10)
        
        fr = tb.Frame(h); fr.pack(side="right")
        
        mode_fr = tb.Labelframe(fr, text="Modo de Análise", padding=(5, 2))
        mode_fr.pack(side="left", padx=10)
        for m in ["F3", "E5", "EXCHANGE", "FULL"]:
            tb.Radiobutton(mode_fr, text=m.capitalize(), value=m, variable=self.analysis_mode_var, bootstyle="toolbutton-outline").pack(side="left", padx=2)
        
        self.pb = tb.Progressbar(fr, mode="indeterminate", length=150, bootstyle="success-striped")
        
        tb.Button(fr, text="Carregar Arquivo", command=self.browse_file, bootstyle="outline").pack(side="left", padx=5)
        self.btn_run = tb.Button(fr, text="▶ Executar", command=self.start_audit, bootstyle="success")
        self.btn_run.pack(side="left", padx=5)
        self.btn_export = tb.Button(fr, text="💾 Exportar", command=self.open_export_modal, state="disabled", bootstyle="info")
        self.btn_export.pack(side="left", padx=5)
        self.btn_cancel = tb.Button(fr, text="✖ Cancelar", command=self.cancel_process, state="disabled", bootstyle="danger")
        self.btn_cancel.pack(side="left", padx=5)

    def _create_tab(self, title, builder):
        f = tb.Frame(self.nb, padding=10)
        self.nb.add(f, text=title)
        builder(f)

    def _build_tab_resumo(self, parent):
        # KPIs
        self.kpi_frame = tb.Frame(parent)
        self.kpi_frame.pack(fill="x", pady=(0,15))
        h = 140
        self._kpis = [
            RoundedCard(self.kpi_frame, "Total Usuários", self.kpi_total, "primary", height=h),
            RoundedCard(self.kpi_frame, "Candidatos", self.kpi_candidates, "danger", height=h),
            RoundedCard(self.kpi_frame, "Conflitos", self.kpi_conflicts, "warning", height=h),
            RoundedCard(self.kpi_frame, "Duplicados (Gr)", self.kpi_dups, "info", height=h)
        ]
        self.kpi_frame.bind("<Configure>", self._layout_kpis)
        
        # Split: Info + Tabela + Log
        split = tb.Frame(parent)
        split.pack(fill="both", expand=True)

        # Esquerda: Interpretação + Tabela
        left_p = tb.Frame(split)
        left_p.pack(side="left", fill="both", expand=True, padx=(0,5))

        info_text = (
            "COMO INTERPRETAR:\n"
            "• Terceiro + Bloqueado: Candidato direto a remoção de todas as licenças.\n"
            f"• Inatividade: Sem troca de senha ou sync há mais de {CFG.inactivity_days_threshold} dias.\n"
            "• Conflito: Possui F3 e E5 simultaneamente.\n"
            "• Modo de Análise: Define sugestões de Downgrade/Padronização."
        )
        lf_info = tb.Labelframe(left_p, text="Guia Rápido", padding=10)
        lf_info.pack(fill="x", pady=(0, 10))
        tb.Label(lf_info, text=info_text, justify="left").pack(anchor="w")

        lf_tab = tb.Labelframe(left_p, text="Métricas da Auditoria", padding=5)
        lf_tab.pack(fill="both", expand=True)
        self.tree_map["Resumo"] = self._create_tree(lf_tab, ["Metrica", "Valor"])

        # Direita: Log
        right = tb.Labelframe(split, text="Log de Processamento", padding=5)
        right.pack(side="right", fill="both", expand=True, padx=(5,0))
        self.txt_log = ScrolledText(right, font=("Consolas", 10), height=10)
        self.txt_log.pack(fill="both", expand=True)

    def _build_tab_candidatos(self, p):
        self._add_search(p)
        self.tree_map["Candidatos"] = self._create_tree(p, ["DisplayName", "UPN", "Severity", "Action", "Rule_Explanation"])

    def _build_tab_padronizacao(self, p):
        self._add_search(p)
        self.tree_map["Padronizacao"] = self._create_tree(p, ["DisplayName", "UPN", "Reasons", "Action", "Rule_Explanation"])

    def _build_tab_conflitos(self, p):
        self._add_search(p)
        self.tree_map["Conflitos"] = self._create_tree(p, ["DisplayName", "UPN", "Action", "Licenses", "Rule_Explanation"])

    def _build_tab_terceiros(self, p):
        self._add_search(p)
        self.tree_map["Terceiros"] = self._create_tree(p, ["DisplayName", "UPN", "Is_Blocked", "Rule_Explanation"])

    def _build_tab_dups_list(self, p):
        self._add_search(p)
        self.tree_map["Duplicados"] = self._create_tree(p, ["dup_score", "keep_recommendation", "DisplayName", "UPN", "dup_notes", "dup_group_id"])

    def _build_tab_dups_group(self, p):
        self.tree_map["Grupos"] = self._create_tree(p, ["ID Grupo", "Qtd", "Score Máx", "UPNs Envolvidos"])

    def _build_tab_inactive(self, p):
        self._add_search(p)
        self.tree_map["Inatividade"] = self._create_tree(p, ["DisplayName", "UPN", "Inactivity_Days", "Inactivity_Message"])

    def _build_tab_all(self, p):
        self._add_search(p)
        self.tree_map["Todos"] = self._create_tree(p, ["DisplayName", "UPN", "Licenses", "Title", "Rule_Explanation"])

    def _add_search(self, parent):
        f = tb.Frame(parent); f.pack(fill="x", pady=(0,5))
        tb.Label(f, text="Filtrar:").pack(side="left")
        e = tb.Entry(f, width=40); e.pack(side="left", padx=5)
        e.bind("<KeyRelease>", lambda ev: self._filter_tree(parent, e.get()))

    def _create_tree(self, parent, cols):
        tv = tb.Treeview(parent, columns=cols, show="headings", bootstyle="primary")
        for c in cols: tv.heading(c, text=c); tv.column(c, width=150)
        
        vsb = tb.Scrollbar(parent, orient="vertical", command=tv.yview)
        tv.configure(yscrollcommand=vsb.set)
        hsb = tb.Scrollbar(parent, orient="horizontal", command=tv.xview)
        tv.configure(xscrollcommand=hsb.set)
        
        tv.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x") 
        
        tv.tag_configure("sev_alta", background="#f9d6d2")
        tv.tag_configure("odd", background="#f8f9fa")
        tv.tag_configure("even", background="#ffffff") # Fix visual
        tv.meta_cols = cols
        return tv

    def _layout_kpis(self, event=None):
        for w in self.kpi_frame.grid_slaves(): w.grid_forget()
        width = self.kpi_frame.winfo_width() or 1
        cols = 4 if width > 1100 else 2
        for c in range(cols): self.kpi_frame.columnconfigure(c, weight=1)
        for i, card in enumerate(self._kpis):
            card.grid(row=i//cols, column=i%cols, sticky="nsew", padx=5, pady=5)
            card.configure(width=200)

    def browse_file(self):
        p = filedialog.askopenfilename()
        if p: self.file_path_var.set(p)

    def start_audit(self):
        if not self.file_path_var.get(): return
        self.is_running = True
        self.toggle_ui(False)
        self.pb.pack(side="left", padx=10); self.pb.start(10)
        if hasattr(self, 'txt_log'): self.txt_log.text.delete("1.0", "end")
        threading.Thread(target=self._run_thread, daemon=True).start()

    def cancel_process(self):
        if self.is_running:
            self.cancel_event.set()
            self.log("Solicitando cancelamento...")

    def _run_thread(self):
        try:
            p = self.file_path_var.get()
            ae = AuditEngine
            mode = self.analysis_mode_var.get()

            df = ae.read_file_robust(p, self.log)
            df = ae.map_columns(df, self.log)
            df = ae.normalize_data(df, self.log)
            df = ae.detect_duplicates(df, self.cancel_event, self.log)
            df = ae.detect_inactivity(df)
            df = ae.analyze_licenses(df)
            df = ae.evaluate_license_profile(df, mode, self.cancel_event, self.log)
            df = ae.apply_global_rules(df, mode, self.cancel_event, self.log)
            df = ae.build_rule_explanations(df, mode) # NEW Explainer
            
            summary_df = ae.build_summary(df) # NEW Summary
            
            self.dict_results = {
                "Resumo": summary_df,
                "Candidatos": df[df["Is_Candidate"]].copy(),
                "Padronizacao": df[df["Reasons"].str.contains("MODO_", na=False)].copy(),
                "Conflitos": df[df["has_conflict_f3_e5"]].copy(),
                "Terceiros": df[df["Is_ThirdParty"]].copy(),
                "Terceiros Bloqueados": df[df["Is_ThirdParty"] & df["Is_Blocked"]].copy(),
                "Duplicados": df[df["dup_strength"] != ""].copy(),
                "Grupos": ae.create_group_view(df[df["dup_strength"] != ""]),
                "Inatividade": df[df["Is_Inactive"]].copy(),
                "Todos": df.copy()
            }
            self.root.after(0, self._finish_audit)
        except Exception as e:
            self.log(f"Erro: {e}")
            self.root.after(0, lambda: [messagebox.showerror("Erro", str(e)), self.toggle_ui(True)])

    def _finish_audit(self):
        self.is_running = False
        self.toggle_ui(True)
        self.pb.stop(); self.pb.pack_forget()
        self.status_var.set("Finalizado")
        
        self.kpi_total.set(str(len(self.dict_results["Todos"])))
        self.kpi_candidates.set(str(len(self.dict_results["Candidatos"])))
        self.kpi_conflicts.set(str(len(self.dict_results["Conflitos"])))
        self.kpi_dups.set(str(len(self.dict_results["Grupos"])))
        
        self._populate_all()

    def _populate_all(self):
        for name, tv in self.tree_map.items():
            self._fill_tree(tv, self.dict_results.get(name, pd.DataFrame()))

    def _fill_tree(self, tv, df):
        tv.delete(*tv.get_children())
        if df.empty: return
        cols = tv.meta_cols
        show = df.head(CFG.preview_limit_step)
        for i, row in enumerate(show.itertuples(index=False)):
            vals = []
            for c in cols:
                if hasattr(row, c): vals.append(getattr(row, c))
                else: vals.append(row._asdict().get(c, ""))
            tag = "even" if i%2==0 else "odd"
            if hasattr(row, "Severity") and getattr(row, "Severity") == "ALTA": tag = "sev_alta"
            tv.insert("", "end", values=vals, tags=(tag,))
        tv.full_df = df

    def _filter_tree(self, parent, query):
        for child in parent.winfo_children():
            if isinstance(child, (ttk.Treeview, tb.Treeview)):
                if not hasattr(child, "full_df"): return
                df = child.full_df
                if query:
                    mask = df.astype(str).apply(lambda x: x.str.lower().str.contains(query.lower(), regex=False)).any(axis=1)
                    df = df[mask]
                self._fill_tree(child, df)
                return

    def open_export_modal(self):
        if not self.dict_results: return
        top = tb.Toplevel(self.root)
        top.title("Opções de Exportação")
        top.geometry("300x450")
        
        tb.Label(top, text="Selecione as abas:", font=("Segoe UI", 10, "bold")).pack(pady=10)
        
        self.export_vars = {}
        default_keys = ["Resumo", "Candidatos", "Padronizacao", "Conflitos", "Duplicados", "Inatividade"]
        
        keys_to_show = [k for k in self.dict_results.keys() if k != "Grupos"]

        for key in keys_to_show:
            var = tk.BooleanVar(value=(key in default_keys))
            self.export_vars[key] = var
            tb.Checkbutton(top, text=key, variable=var).pack(anchor="w", padx=40, pady=2)
            
        tb.Button(top, text="Gerar Excel", bootstyle="success", command=lambda: [self.run_export(top)]).pack(fill="x", padx=20, pady=20)

    def run_export(self, modal):
        modal.destroy()
        path = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Audit_{datetime.now().strftime('%Y%m%d')}.xlsx")
        if not path: return
        
        try:
            with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
                wb = writer.book
                fmt = wb.add_format({'bold': True, 'bg_color': '#D9E1F2', 'border': 1})
                
                for name, df in self.dict_results.items():
                    if name not in self.export_vars or not self.export_vars[name].get(): continue
                    if df.empty: continue
                    
                    drop = ["RowId", "dup_score"]
                    out = df.drop(columns=[c for c in drop if c in df.columns], errors='ignore')
                    
                    out.to_excel(writer, sheet_name=name[:30], index=False)
                    if len(out) > 0:
                        ws = writer.sheets[name[:30]]
                        ws.autofilter(0, 0, len(out), len(out.columns)-1)
                        for i, col in enumerate(out.columns):
                            ws.write(0, i, col, fmt)
                            ws.set_column(i, i, 20)
                            
            messagebox.showinfo("Sucesso", "Exportado com sucesso!")
            safe_open_folder(path)
        except Exception as e:
            messagebox.showerror("Erro", str(e))

    # --- Persistence ---
    def load_settings(self):
        if os.path.exists(SETTINGS_FILE):
            try:
                with open(SETTINGS_FILE, 'r') as f:
                    data = json.load(f)
                    self.analysis_mode_var.set(data.get("mode", "FULL"))
                    path = data.get("last_file", "")
                    if path and os.path.exists(path): self.file_path_var.set(path)
            except: pass

    def on_close(self):
        data = {
            "mode": self.analysis_mode_var.get(),
            "last_file": self.file_path_var.get()
        }
        try:
            with open(SETTINGS_FILE, 'w') as f: json.dump(data, f)
        except: pass
        self.root.destroy()

    def log(self, msg): self.log_queue.put(msg)
    def _process_logs(self):
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            if hasattr(self, 'txt_log'): self.txt_log.text.insert("end", f"{msg}\n"); self.txt_log.text.see("end")
        self.root.after(100, self._process_logs)
    def toggle_ui(self, enable):
        st = "normal" if enable else "disabled"
        self.btn_run.config(state=st)
        self.btn_export.config(state=st)
        self.btn_cancel.config(state="normal" if not enable else "disabled")

if __name__ == "__main__":
    root = tb.Window(themename="flatly")
    app = AuditorApp(root)
    root.mainloop()