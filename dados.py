"""
Aegis Analytics Engine - Plataforma Autónoma de Engenharia de Dados
Módulo de nível industrial focado em profiling local, limpeza estatística,
visualização gráfica integrada e exportação de relatórios ricos em HTML com gráficos inline.
"""

import os
import json
import math
import argparse
import threading
import base64
from io import BytesIO
from datetime import datetime
from typing import List, Dict, Any, Union, Tuple, Optional
import pandas as pd
import numpy as np

# Importações nativas para suporte a Interface Gráfica (GUI)
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# Importação condicional para visualização científica
try:
    import matplotlib
    matplotlib.use('Agg')  # Backend sem GUI para uso em threads
    import matplotlib.pyplot as plt
    from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False


class AegisAnalyticsEngine:
    def __init__(self, log_callback=None):
        """
        Inicializa o motor Aegis Analytics em modo 100% local e offline.
        
        Args:
            log_callback (callable): Função opcional para redirecionar logs (útil para a GUI).
        """
        self.dataset: Optional[pd.DataFrame] = None
        self.raw_dataset: Optional[pd.DataFrame] = None
        self.schema: Dict[str, Dict[str, Any]] = {}
        self.correlation_matrix: Optional[pd.DataFrame] = None
        self.logs: List[str] = []
        self.log_callback = log_callback
        self.clean_report: Dict[str, Any] = {
            "imputations_count": 0,
            "outliers_detected": 0,
            "filename": ""
        }
        self.figures: Dict[str, plt.Figure] = {}  # Guarda as figuras geradas
        
        self._write_log("Aegis Analytics Engine inicializado em modo puramente estatístico (local).")
        if not HAS_MATPLOTLIB:
            self._write_log("Aviso: 'matplotlib' não encontrado. Gráficos desativados. Instale com 'pip install matplotlib'.", "WARNING")

    def _write_log(self, message: str, level: str = "INFO") -> None:
        """Regista mensagens operacionais internamente com carimbo de data/hora."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = f"[{timestamp}] [{level}] {message}"
        self.logs.append(log_entry)
        
        # Envia para o terminal
        print(log_entry)
        
        # Envia para a interface gráfica se o callback estiver registado
        if self.log_callback:
            self.log_callback(log_entry)

    def load_data(self, filepath: str) -> pd.DataFrame:
        """
        Carrega dados heterogéneos de forma tolerante a falhas. Suporta CSV, JSON e Excel (.xlsx, .xls).
        """
        self._write_log(f"Iniciando ingestão de dados de: {filepath}")
        if not os.path.exists(filepath):
            self._write_log(f"Ficheiro não encontrado: {filepath}", "ERROR")
            raise FileNotFoundError(f"Caminho inválido: {filepath}")

        try:
            if filepath.endswith('.csv'):
                self.raw_dataset = pd.read_csv(filepath, low_memory=False)
            elif filepath.endswith('.json'):
                self.raw_dataset = pd.read_json(filepath)
            elif filepath.endswith(('.xlsx', '.xls')):
                try:
                    import openpyxl
                except ImportError:
                    raise ImportError(
                        "A biblioteca 'openpyxl' é necessária para ler ficheiros Excel.\n"
                        "Por favor, instale-a executando: pip install openpyxl"
                    )
                self.raw_dataset = pd.read_excel(filepath)
            else:
                raise ValueError("Formato não suportado. Utilize apenas ficheiros .csv, .json, .xlsx ou .xls.")
            
            self.dataset = self.raw_dataset.copy()
            self.clean_report["filename"] = os.path.basename(filepath)
            self._write_log(f"Ingestão concluída. Dimensões: {self.dataset.shape[0]} linhas x {self.dataset.shape[1]} colunas", "SUCCESS")
            return self.dataset
        except Exception as e:
            self._write_log(f"Falha crítica no parsing do ficheiro: {str(e)}", "ERROR")
            raise

    def import_dataframe(self, df: pd.DataFrame, name: str = "dataframe_memoria") -> pd.DataFrame:
        """Permite importar diretamente um DataFrame do Pandas em memória."""
        if not isinstance(df, pd.DataFrame):
            self._write_log("Falha ao importar: O objeto fornecido não é um DataFrame válido.", "ERROR")
            raise TypeError("O objeto fornecido deve ser um pd.DataFrame.")
        
        self._write_log(f"A importar DataFrame externo em memória: {name}")
        self.raw_dataset = df.copy()
        self.dataset = self.raw_dataset.copy()
        self.clean_report["filename"] = name
        self._write_log(f"Importação concluída. Dimensões: {self.dataset.shape[0]} linhas x {self.dataset.shape[1]} colunas", "SUCCESS")
        return self.dataset

    def load_demo_data(self) -> pd.DataFrame:
        """Gera e carrega um dataset corporativo de teste altamente instável (com nulos e outliers)."""
        self._write_log("A carregar massa de dados sintética de teste corporativo...")
        np.random.seed(42)
        n_rows = 400
        
        segments = ["Financeiro", "Logística", "E-Commerce", "Saúde", "Setor Público"]
        regions = ["Sul", "Sudeste", "Nordeste", "Norte", "Centro-Oeste"]
        
        data = {
            "ID_Transacao": [f"TX-{1000 + i}" for i in range(n_rows)],
            "Data_Registro": pd.date_range(start="2026-01-01", periods=n_rows, freq="h").strftime("%Y-%m-%d").tolist(),
            "Segmento_Negocio": [np.random.choice(segments) if i % 15 != 0 else None for i in range(n_rows)],
            "Regiao_Operacional": [np.random.choice(regions) for _ in range(n_rows)],
            "Valor_Contrato": [float(np.random.randint(10000, 85000)) if i % 35 != 0 else float(np.random.randint(150000, 450000)) for i in range(n_rows)],
            "Margem_Lucro": [float(np.random.uniform(0.08, 0.50)) if i % 20 != 0 else None for i in range(n_rows)],
            "Lead_Rating": [int(np.random.randint(1, 6)) for _ in range(n_rows)]
        }
        
        self.raw_dataset = pd.DataFrame(data)
        self.dataset = self.raw_dataset.copy()
        self.clean_report["filename"] = "vendas_corporativas_demo.xlsx"
        self._write_log(f"Massa sintética injetada. Dimensões: {self.dataset.shape[0]} linhas x {self.dataset.shape[1]} colunas", "SUCCESS")
        return self.dataset

    def run_pipeline(self, impute_nulls: bool = True, detect_outliers: bool = True, infer_types: bool = True) -> Dict[str, Any]:
        """Executa sequencialmente o pipeline analítico estatístico local."""
        if self.dataset is None:
            raise ValueError("Nenhum dado ativo carregado no motor.")

        self._write_log("Iniciando pipeline analítico autônomo local...")
        columns = self.dataset.columns.tolist()
        self.schema = {}
        self.clean_report["imputations_count"] = 0
        self.clean_report["outliers_detected"] = 0

        # 1. Inferência de Tipos Semânticos e Profiling Primitivo
        for col in columns:
            null_count = int(self.dataset[col].isnull().sum())
            distinct_count = int(self.dataset[col].nunique())
            pd_type = str(self.dataset[col].dtype)
            inferred_type = 'categorical'

            if 'int' in pd_type or 'float' in pd_type:
                inferred_type = 'numeric'
            elif 'bool' in pd_type:
                inferred_type = 'boolean'
            
            if infer_types:
                if inferred_type == 'numeric' and distinct_count <= 5 and distinct_count > 0:
                    inferred_type = 'categorical'
                
                if inferred_type == 'categorical':
                    non_nulls = self.dataset[col].dropna()
                    if not non_nulls.empty:
                        first_val = non_nulls.iloc[0]
                        if isinstance(first_val, str) and self._is_date_string(first_val):
                            inferred_type = 'temporal'
                        
                geo_keywords = ['brasil', 'sp', 'rj', 'estado', 'cidade', 'regiao', 'pais', 'uf', 'bairro', 'região', 'país']
                if inferred_type == 'categorical' and any(keyword in col.lower() for keyword in geo_keywords):
                    inferred_type = 'geographical'

            self.schema[col] = {
                "type": inferred_type,
                "null_count": null_count,
                "distinct_count": distinct_count,
                "num_stats": {},
                "outliers_count": 0,
                "top_categories": []
            }

        # 2. Imputação de Valores Nulos
        if impute_nulls:
            for col in columns:
                s = self.schema[col]
                if s["null_count"] > 0:
                    if s["type"] == 'numeric':
                        impute_val = float(self.dataset[col].median())
                        self.dataset[col] = self.dataset[col].fillna(impute_val)
                    else:
                        mode_series = self.dataset[col].mode()
                        impute_val = mode_series.iloc[0] if not mode_series.empty else "DESCONHECIDO"
                        self.dataset[col] = self.dataset[col].fillna(impute_val)
                    
                    self.clean_report["imputations_count"] += s["null_count"]
                    self._write_log(f"Coluna [{col}] normalizada. {s['null_count']} nulos imputados com [{impute_val}].", "WARNING")

        # 3. Varredura Estatística Avançada & Outliers (IQR Method)
        for col in columns:
            s = self.schema[col]
            if s["type"] == 'numeric':
                series_clean = self.dataset[col].dropna()
                if not series_clean.empty:
                    q1 = float(series_clean.quantile(0.25))
                    q3 = float(series_clean.quantile(0.75))
                    iqr = q3 - q1
                    lower_bound = q1 - 1.5 * iqr
                    upper_bound = q3 + 1.5 * iqr

                    outliers = series_clean[(series_clean < lower_bound) | (series_clean > upper_bound)]
                    s["outliers_count"] = len(outliers)
                    self.clean_report["outliers_detected"] += len(outliers)

                    s["num_stats"] = {
                        "min": float(series_clean.min()),
                        "max": float(series_clean.max()),
                        "mean": float(series_clean.mean()),
                        "median": float(series_clean.median()),
                        "std": float(series_clean.std())
                    }
            elif s["type"] in ['categorical', 'geographical']:
                top_cats = self.dataset[col].value_counts().head(5).index.tolist()
                s["top_categories"] = top_cats

        # 4. Cálculo de Correlações Lineares de Pearson
        numeric_cols = [col for col, meta in self.schema.items() if meta["type"] == 'numeric']
        if len(numeric_cols) >= 2:
            self.correlation_matrix = self.dataset[numeric_cols].corr(method='pearson')
            self._write_log("Matriz de correlação Pearson calculada com sucesso.", "SUCCESS")
        
        # 5. Geração Automática dos Gráficos em Memória
        if HAS_MATPLOTLIB:
            self._generate_plots()

        self._write_log("Pipeline analítico local executado com sucesso.", "SUCCESS")
        return {
            "schema": self.schema,
            "clean_report": self.clean_report,
            "correlation_matrix": self.correlation_matrix.to_dict() if self.correlation_matrix is not None else None
        }

    def _is_date_string(self, val: str) -> bool:
        try:
            pd.to_datetime(val)
            return True
        except (ValueError, TypeError):
            return False

    def _generate_plots(self):
        """Gera e armazena os gráficos estatísticos do dataset em memória."""
        self._write_log("Gerando visualizações gráficas estatísticas...")
        
        # Resetar dicionário de figuras
        self.figures.clear()
        plt.close('all')

        numeric_cols = [col for col, meta in self.schema.items() if meta["type"] == 'numeric']
        
        # 1. Gráfico de Distribuição da variável numérica de maior amplitude/escala
        if numeric_cols:
            best_col = max(numeric_cols, key=lambda c: self.schema[c]["num_stats"].get("max", 0) - self.schema[c]["num_stats"].get("min", 0))
            fig, ax = plt.subplots(figsize=(6, 4.2), facecolor='#1e293b')
            ax.set_facecolor('#0f172a')
            
            data_points = self.dataset[best_col].dropna()
            ax.hist(data_points, bins=25, color='#10b981', edgecolor='#059669', alpha=0.8, density=True)
            
            # Adiciona linha de densidade aproximada (KDE suave)
            try:
                from scipy.stats import gaussian_kde
                xs = np.linspace(data_points.min(), data_points.max(), 200)
                kde = gaussian_kde(data_points)
                ax.plot(xs, kde(xs), color='#67e8f9', linewidth=2, label="Densidade (KDE)")
            except ImportError:
                pass # Caso não haja scipy, renderiza apenas o histograma estrito

            ax.set_title(f"Distribuição de Frequência: {best_col}", color='#f1f5f9', fontsize=11, fontweight='bold')
            ax.tick_params(colors='#94a3b8', labelsize=8)
            for spine in ax.spines.values():
                spine.set_color('#334155')
            ax.grid(True, color='#1e293b', linestyle='--', alpha=0.5)
            self.figures["dist_chart"] = fig

        # 2. Heatmap da Matriz de Correlação
        if self.correlation_matrix is not None and len(numeric_cols) >= 2:
            fig, ax = plt.subplots(figsize=(6, 4.2), facecolor='#1e293b')
            ax.set_facecolor('#0f172a')
            
            corr = self.correlation_matrix.to_numpy()
            im = ax.imshow(corr, cmap='coolwarm', vmin=-1, vmax=1)
            
            # Adiciona labels nos eixos
            ax.set_xticks(np.arange(len(numeric_cols)))
            ax.set_yticks(np.arange(len(numeric_cols)))
            ax.set_xticklabels(numeric_cols, rotation=45, ha="right", color='#94a3b8', fontsize=8)
            ax.set_yticklabels(numeric_cols, color='#94a3b8', fontsize=8)
            
            # Inserir os coeficientes textuais em cada célula do Heatmap
            for i in range(len(numeric_cols)):
                for j in range(len(numeric_cols)):
                    ax.text(j, i, f"{corr[i, j]:.2f}", ha="center", va="center", 
                            color="black" if abs(corr[i, j]) > 0.5 else "#f1f5f9", fontsize=8, fontweight='bold')
            
            ax.set_title("Matriz de Correlação de Pearson (R)", color='#f1f5f9', fontsize=11, fontweight='bold')
            for spine in ax.spines.values():
                spine.set_color('#334155')
            
            # Barra de cores adaptada
            cbar = fig.colorbar(im, ax=ax, shrink=0.8)
            cbar.ax.yaxis.set_tick_params(color='#94a3b8', labelcolor='#94a3b8', labelsize=8)
            cbar.outline.set_edgecolor('#334155')

            self.figures["corr_chart"] = fig

        # 3. Boxplot de Dispersão e Outliers
        if len(numeric_cols) > 0:
            fig, ax = plt.subplots(figsize=(6, 4.2), facecolor='#1e293b')
            ax.set_facecolor('#0f172a')
            
            # Normalização simples (MinMax) para colocar múltiplas numéricas na mesma escala do boxplot
            norm_data = []
            valid_cols = []
            for col in numeric_cols[:4]:  # Limita a 4 colunas para não poluir
                col_data = self.dataset[col].dropna()
                if col_data.max() != col_data.min():
                    norm_series = (col_data - col_data.min()) / (col_data.max() - col_data.min())
                    norm_data.append(norm_series)
                    valid_cols.append(col)
            
            if norm_data:
                bp = ax.boxplot(norm_data, patch_artist=True, tick_labels=valid_cols)
                
                # Customização visual do Boxplot
                for box in bp['boxes']:
                    box.set(facecolor=(16/255, 185/255, 129/255, 0.4), color='#10b981', linewidth=1.5)
                for median in bp['medians']:
                    median.set(color='#f43f5e', linewidth=2)
                for flier in bp['fliers']:
                    flier.set(marker='o', markerfacecolor='#f59e0b', markeredgecolor='#d97706', alpha=0.6, markersize=5)
                
                ax.set_title("Análise de Dispersão e Outliers (MinMax Normalizado)", color='#f1f5f9', fontsize=11, fontweight='bold')
                ax.tick_params(colors='#94a3b8', labelsize=8)
                ax.set_xticklabels(valid_cols, rotation=15, color='#94a3b8', fontsize=8)
                for spine in ax.spines.values():
                    spine.set_color('#334155')
                ax.grid(True, color='#1e293b', linestyle='--', alpha=0.5)
                self.figures["outlier_chart"] = fig

    def generate_local_insights(self) -> Dict[str, Any]:
        """Gera o relatório analítico local baseado puramente em regras estatísticas."""
        insights = []
        recomendacoes = [
            "Garantir a verificação constante de consistência operacional dos dados na fronteira de ingestão.",
            "Utilizar algoritmos de regressão local e árvores de decisão para colunas com correlações superiores a 0.70."
        ]
        
        if self.correlation_matrix is not None:
            cols = self.correlation_matrix.columns
            for i in range(len(cols)):
                for j in range(i + 1, len(cols)):
                    val = self.correlation_matrix.iloc[i, j]
                    col_x, col_y = cols[i], cols[j]
                    if val > 0.75:
                        insights.append({
                            "tipo": "Correlação Direta Forte",
                            "descricao": f"Relação linear forte de ({val:.3f}) entre as variáveis [{col_x}] e [{col_y}].",
                            "impacto": "Flutuações de um atributo refletem impactos diretos nas métricas do outro simultaneamente."
                        })
                    elif val < -0.75:
                        insights.append({
                            "tipo": "Correlação Inversa Forte",
                            "descricao": f"Incompatibilidade linear severa ({val:.3f}) identificada entre as colunas [{col_x}] e [{col_y}].",
                            "impacto": "O crescimento de um fator impõe compressão proporcional à atividade do outro."
                        })

        for col, meta in self.schema.items():
            if meta["type"] == 'numeric' and meta["num_stats"]:
                stats = meta["num_stats"]
                if stats["mean"] != 0:
                    cv = stats["std"] / stats["mean"]
                    if cv > 0.6:
                        insights.append({
                            "tipo": "Instabilidade / Alta Volatilidade",
                            "descricao": f"A coluna contínua [{col}] apresenta desvio padrão elevado em relação à média (CV: {cv:.2f}).",
                            "impacto": "Sinais de dados altamente sensíveis a anomalias temporárias ou flutuações sazonais abruptas."
                        })
                        recomendacoes.append(f"Considerar normalização em log ou escalonamento MinMax para o vetor [{col}] antes de modelos de ML.")

        total_rows = len(self.dataset) if self.dataset is not None else 1
        for col, meta in self.schema.items():
            if meta["type"] == 'numeric' and meta["outliers_count"] > (total_rows * 0.05):
                insights.append({
                    "tipo": "Anomalias de Cauda Larga",
                    "descricao": f"O vetor numérico [{col}] comporta {meta['outliers_count']} outliers identificados no pipeline IQR.",
                    "impacto": "Elevada probabilidade de transações atípicas, erros de captura física ou ruídos sistémicos de infraestrutura."
                })
                recomendacoes.append(f"Aplicar técnicas de Winsorization ou remoção truncada de outliers no vetor [{col}] se o objetivo for linearidade.")

        if not insights:
            insights.append({
                "tipo": "Padrão Linear Homogêneo",
                "descricao": "Os vetores estatísticos do dataset comportam-se dentro de faixas normais com variabilidade balanceada.",
                "impacto": "Modelagem clássica de inferência estatística pode ser aplicada diretamente sem grandes transformações."
            })

        return {
            "titulo": "Diagnóstico Analítico de Integridade Estatística (Offline)",
            "resumo_executivo": f"Processamento de dados locais concluído com sucesso. Ingestão efetuada com {total_rows} registros estruturados.",
            "insights": insights,
            "recomendacoes_tecnicas": list(set(recomendacoes))
        }

    def export_html_report(self, output_path: str, local_report: Dict[str, Any]) -> None:
        """
        Gera um relatório profissional executivo em HTML autossuficiente (Inline)
        contendo todo o CSS estruturado com Tailwind e os gráficos convertidos para Base64.
        """
        self._write_log(f"Iniciando exportação de relatório executivo rico em HTML para: {output_path}")
        
        # Converte figuras do Matplotlib para strings em formato Base64
        charts_html = ""
        for name, fig in self.figures.items():
            buf = BytesIO()
            fig.savefig(buf, format='png', dpi=120, facecolor='#1e293b', bbox_inches='tight')
            buf.seek(0)
            img_b64 = base64.b64encode(buf.read()).decode('utf-8')
            
            title = "Visualização Analítica"
            if name == "dist_chart": title = "Histograma de Frequência"
            elif name == "corr_chart": title = "Mapa de Correlações Lineares"
            elif name == "outlier_chart": title = "Dispersão e Detecção de Outliers"
            
            charts_html += f"""
            <div class="bg-slate-800/40 border border-slate-700 rounded-xl p-5 flex flex-col items-center">
                <h3 class="text-sm font-semibold text-slate-300 font-mono mb-3 uppercase tracking-wider">{title}</h3>
                <img src="data:image/png;base64,{img_b64}" class="w-full h-auto max-w-lg rounded-lg shadow-xl border border-slate-900" />
            </div>
            """

        insights_html = ""
        for idx, insight in enumerate(local_report.get('insights', []), 1):
            insights_html += f"""
            <div class="p-4 bg-slate-900/60 border-l-4 border-emerald-500 rounded-r-lg">
                <div class="flex items-center gap-2">
                    <span class="text-[10px] font-mono font-bold uppercase tracking-widest px-2 py-0.5 rounded bg-emerald-500/10 text-emerald-400 border border-emerald-500/20">#{idx} - {insight.get('tipo')}</span>
                </div>
                <p class="text-xs text-slate-200 mt-2 font-mono">{insight.get('descricao')}</p>
                <p class="text-xs text-slate-400 italic mt-1">Impacto: {insight.get('impacto')}</p>
            </div>
            """

        recs_html = ""
        for rec in local_report.get('recomendacoes_tecnicas', []):
            recs_html += f"""<li class="text-xs text-slate-300 font-mono list-disc list-inside">{rec}</li>"""

        dictionary_rows = ""
        for col, meta in self.schema.items():
            dictionary_rows += f"""
            <tr class="hover:bg-slate-800/30 border-b border-slate-800/80 transition-colors">
                <td class="p-3 text-emerald-400 font-mono font-bold">{col}</td>
                <td class="p-3"><span class="px-2 py-0.5 text-[10px] bg-slate-800 text-slate-300 rounded border border-slate-700">{meta['type'].upper()}</span></td>
                <td class="p-3 text-slate-300 font-mono">{meta['distinct_count']}</td>
                <td class="p-3 text-slate-300 font-mono">{meta['null_count']}</td>
                <td class="p-3 text-amber-400 font-mono font-bold">{meta['outliers_count']}</td>
            </tr>
            """

        html_content = f"""<!DOCTYPE html>
<html lang="pt-BR" class="h-full bg-slate-950 text-slate-100">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Relatório Executivo Aegis Engine</title>
    <script src="https://cdn.tailwindcss.com"></script>
</head>
<body class="bg-slate-950 text-slate-100 font-sans p-8">
    <div class="max-w-5xl mx-auto space-y-8">
        
        <!-- HEADER -->
        <header class="flex flex-col md:flex-row md:items-center md:justify-between border-b border-slate-800 pb-6">
            <div>
                <h1 class="text-2xl font-black tracking-wider text-emerald-400 font-mono">AEGIS CORE REPORT</h1>
                <p class="text-xs text-slate-400 font-mono mt-1 uppercase">Relatório Técnico Autónomo Estatístico de Ingestão</p>
            </div>
            <div class="text-right mt-4 md:mt-0 font-mono text-xs text-slate-500">
                <p>Gerado em: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</p>
                <p>Status: 100% Estatístico Offline</p>
            </div>
        </header>

        <!-- SUMÁRIO OPERACIONAL -->
        <section class="grid grid-cols-1 md:grid-cols-4 gap-4">
            <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <p class="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Dimensões do Dataset</p>
                <p class="text-lg font-bold text-slate-200 mt-1">{len(self.dataset) if self.dataset is not None else 0} x {len(self.schema)}</p>
            </div>
            <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <p class="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Integridade Geral</p>
                <p class="text-lg font-bold text-slate-200 mt-1">{((1 - (sum(meta['null_count'] for meta in self.schema.values()) / (len(self.dataset) * len(self.schema) if self.dataset is not None else 1))) * 100):.2f}%</p>
            </div>
            <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <p class="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Imputações Efetuadas</p>
                <p class="text-lg font-bold text-cyan-400 mt-1">{self.clean_report['imputations_count']}</p>
            </div>
            <div class="bg-slate-900 border border-slate-800 rounded-xl p-5">
                <p class="text-[10px] font-mono text-slate-500 uppercase tracking-widest">Outliers Identificados</p>
                <p class="text-lg font-bold text-amber-500 mt-1">{self.clean_report['outliers_detected']}</p>
            </div>
        </section>

        <!-- DIAGNÓSTICO -->
        <section class="bg-slate-900 border border-slate-800 rounded-xl p-6 space-y-4">
            <h2 class="text-base font-bold font-mono text-slate-200 uppercase tracking-widest border-b border-slate-800 pb-2">Diagnóstico de Engenharia de Dados</h2>
            <div>
                <h4 class="text-sm font-bold text-emerald-400 font-mono">{local_report.get('titulo')}</h4>
                <p class="text-xs text-slate-300 mt-1 leading-relaxed font-mono italic">"{local_report.get('resumo_executivo')}"</p>
            </div>
            <div class="grid grid-cols-1 md:grid-cols-2 gap-4 mt-4">
                <div class="space-y-3">
                    <h3 class="text-xs font-bold text-slate-400 font-mono uppercase tracking-wider">Insights Estatísticos Críticos:</h3>
                    {insights_html}
                </div>
                <div class="space-y-3 bg-slate-950/40 p-4 border border-slate-800 rounded-lg">
                    <h3 class="text-xs font-bold text-slate-400 font-mono uppercase tracking-wider">Ações de Engenharia Sugeridas:</h3>
                    <ul class="space-y-2 pl-2">
                        {recs_html}
                    </ul>
                </div>
            </div>
        </section>

        <!-- VISUALIZAÇÕES -->
        <section class="space-y-4">
            <h2 class="text-base font-bold font-mono text-slate-200 uppercase tracking-widest border-b border-slate-800 pb-2">Visualizações Analíticas Integradas</h2>
            <div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
                {charts_html}
            </div>
        </section>

        <!-- DICIONÁRIO DE DADOS -->
        <section class="space-y-4">
            <h2 class="text-base font-bold font-mono text-slate-200 uppercase tracking-widest border-b border-slate-800 pb-2">Dicionário de Variáveis & Sanidade</h2>
            <div class="overflow-hidden border border-slate-800 rounded-xl bg-slate-900/40">
                <table class="w-full text-left text-xs">
                    <thead class="bg-slate-900 border-b border-slate-800 font-mono text-slate-400 uppercase tracking-wider">
                        <tr>
                            <th class="p-3">Coluna</th>
                            <th class="p-3">Tipo Inferido</th>
                            <th class="p-3">Valores Únicos</th>
                            <th class="p-3">Valores Nulos</th>
                            <th class="p-3">Outliers</th>
                        </tr>
                    </thead>
                    <tbody class="divide-y divide-slate-800/60 text-slate-300">
                        {dictionary_rows}
                    </tbody>
                </table>
            </div>
        </section>

    </div>
</body>
</html>
"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html_content)
        self._write_log("Exportação do relatório interativo HTML concluída com sucesso.", "SUCCESS")


class AegisAppGUI:
    """Interface Gráfica Nativa de Alta Fidelidade para o Aegis Analytics."""
    def __init__(self, root):
        self.root = root
        self.root.title("Aegis Analytics Engine - Desktop Terminal")
        self.root.geometry("900x700")
        self.root.configure(bg="#0f172a")  # Slate-900 escuro
        
        self.filepath = ""
        self.report_data = None
        self.chart_canvases: List[FigureCanvasTkAgg] = []
        
        self.setup_styles()
        self.build_ui()
        
        # O motor é intencionalmente inicializado APÓS a UI e seus widgets estarem totalmente disponíveis
        self.engine = AegisAnalyticsEngine(log_callback=self.update_log_display)

    def setup_styles(self):
        """Define o visual moderno escuro no Tkinter."""
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Cores do Sistema
        self.style.configure('.', background='#0f172a', foreground='#f1f5f9')
        self.style.configure('TFrame', background='#0f172a')
        self.style.configure('Card.TFrame', background='#1e293b', borderwidth=1, relief='solid')
        self.style.configure('FlatCard.TFrame', background='#1e293b', borderwidth=0, relief='flat')
        
        # Estilo dos Botões
        self.style.configure('Primary.TButton', background='#10b981', foreground='#030712', borderwidth=0, font=('Helvetica', 10, 'bold'))
        self.style.map('Primary.TButton', background=[('active', '#34d399')])
        
        self.style.configure('Secondary.TButton', background='#334155', foreground='#f1f5f9', borderwidth=0, font=('Helvetica', 10))
        self.style.map('Secondary.TButton', background=[('active', '#475569')])
        
        # Notebook / Tabs
        self.style.configure('TNotebook', background='#0f172a', borderwidth=0)
        self.style.configure('TNotebook.Tab', background='#1e293b', foreground='#94a3b8', font=('Helvetica', 9, 'bold'), padding=(15, 5))
        self.style.map('TNotebook.Tab', background=[('selected', '#10b981')], foreground=[('selected', '#030712')])
        
        # Label de Título
        self.style.configure('Header.TLabel', background='#0f172a', foreground='#10b981', font=('Helvetica', 16, 'bold'))
        self.style.configure('SubHeader.TLabel', background='#0f172a', foreground='#94a3b8', font=('Helvetica', 9, 'italic'))
        self.style.configure('LabelCard.TLabel', background='#1e293b', foreground='#f1f5f9', font=('Helvetica', 10, 'bold'))
        self.style.configure('Status.TLabel', background='#1e293b', foreground='#10b981', font=('Courier', 10, 'bold'))

    def build_ui(self):
        """Desenha a árvore de componentes visuais reorganizada em Abas."""
        # Top Header Frame
        header_frame = ttk.Frame(self.root, padding=20)
        header_frame.pack(fill='x')
        
        title_label = ttk.Label(header_frame, text="AEGIS ANALYTICS - CONTROL PANEL", style='Header.TLabel')
        title_label.pack(anchor='w')
        subtitle_label = ttk.Label(header_frame, text="SISTEMA PROFISSIONAL DE ENGENHARIA DE DADOS E RELATÓRIOS OFFLINE", style='SubHeader.TLabel')
        subtitle_label.pack(anchor='w', pady=(2, 0))
        
        # Notebook de Abas Principal
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        # ABA 1: Painel de Controle e Ingestão
        tab_control = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(tab_control, text="Painel de Controle")
        
        # ABA 2: Visualizações de Gráficos Integrados
        self.tab_visuals = ttk.Frame(self.notebook, padding=15)
        self.notebook.add(self.tab_visuals, text="Visualizações e Gráficos")

        # --- CONTEÚDO DA ABA 1 (PAINEL DE CONTROLE) ---
        # CARD 1: File Ingestion
        ingest_card = ttk.Frame(tab_control, style='Card.TFrame', padding=15)
        ingest_card.pack(fill='x', pady=(0, 15))
        
        card1_title = ttk.Label(ingest_card, text="1. Ingestão de Planilhas e Datasets", style='LabelCard.TLabel')
        card1_title.pack(anchor='w', pady=(0, 10))
        
        file_input_frame = ttk.Frame(ingest_card, style='FlatCard.TFrame')
        file_input_frame.pack(fill='x')
        
        self.file_path_var = tk.StringVar(value="Selecione um ficheiro (.csv, .json, .xlsx)...")
        file_entry = tk.Entry(file_input_frame, textvariable=self.file_path_var, bg="#0f172a", fg="#94a3b8", insertbackground="white", relief="flat", font=('Helvetica', 10))
        file_entry.pack(side='left', fill='x', expand=True, padx=(0, 10), ipady=5)
        
        btn_browse = ttk.Button(file_input_frame, text="Procurar", style='Secondary.TButton', command=self.browse_file)
        btn_browse.pack(side='left', padx=(0, 5))
        
        btn_demo = ttk.Button(file_input_frame, text="Carregar Demo", style='Secondary.TButton', command=self.load_demo)
        btn_demo.pack(side='left')

        # CARD 2: Pipeline Configuration
        pipeline_card = ttk.Frame(tab_control, style='Card.TFrame', padding=15)
        pipeline_card.pack(fill='x', pady=(0, 15))
        
        card2_title = ttk.Label(pipeline_card, text="2. Configuração de Pipeline de Dados", style='LabelCard.TLabel')
        card2_title.pack(anchor='w', pady=(0, 10))
        
        chk_frame = ttk.Frame(pipeline_card, style='FlatCard.TFrame')
        chk_frame.pack(fill='x')
        
        self.opt_impute = tk.BooleanVar(value=True)
        self.opt_outliers = tk.BooleanVar(value=True)
        self.opt_infer = tk.BooleanVar(value=True)
        
        c1 = tk.Checkbutton(chk_frame, text="Imputação de Valores Nulos", variable=self.opt_impute, bg='#1e293b', fg='#f1f5f9', activebackground='#1e293b', selectcolor='#0f172a')
        c1.pack(side='left', padx=(0, 20))
        
        c2 = tk.Checkbutton(chk_frame, text="Mapear Outliers (IQR)", variable=self.opt_outliers, bg='#1e293b', fg='#f1f5f9', activebackground='#1e293b', selectcolor='#0f172a')
        c2.pack(side='left', padx=(0, 20))
        
        c3 = tk.Checkbutton(chk_frame, text="Inferência Semântica de Tipos", variable=self.opt_infer, bg='#1e293b', fg='#f1f5f9', activebackground='#1e293b', selectcolor='#0f172a')
        c3.pack(side='left')

        # CARD 3: Actions & Progress Indicators
        actions_card = ttk.Frame(tab_control, style='Card.TFrame', padding=15)
        actions_card.pack(fill='x', pady=(0, 15))
        
        self.btn_run = ttk.Button(actions_card, text="Executar Pipeline Analítico", style='Primary.TButton', command=self.run_analysis_thread)
        self.btn_run.pack(side='left', padx=(0, 15))
        
        self.btn_export = ttk.Button(actions_card, text="Exportar Relatório Rico HTML", style='Secondary.TButton', state='disabled', command=self.export_report)
        self.btn_export.pack(side='left')
        
        self.status_var = tk.StringVar(value="AGUARDANDO INGESTÃO")
        status_label = ttk.Label(actions_card, textvariable=self.status_var, style='Status.TLabel')
        status_label.pack(side='right', padx=10)

        # LOGGING TERMINAL CONSOLE
        console_title = ttk.Label(tab_control, text="Terminal Integrado de Execução (Logs)", style='SubHeader.TLabel')
        console_title.pack(anchor='w', pady=(5, 5))
        
        console_frame = ttk.Frame(tab_control, style='Card.TFrame', padding=2)
        console_frame.pack(fill='both', expand=True)
        
        self.log_text = tk.Text(console_frame, bg="#030712", fg="#34d399", insertbackground="white", relief="flat", font=('Courier', 9))
        self.log_text.pack(fill='both', expand=True, side='left')
        
        scrollbar = ttk.Scrollbar(console_frame, command=self.log_text.yview)
        scrollbar.pack(fill='y', side='right')
        self.log_text.config(yscrollcommand=scrollbar.set)

        # --- CONTEÚDO DA ABA 2 (VISUALIZAÇÕES) ---
        # Inicializado como uma mensagem instrucional se não houver gráficos ativos
        self.visuals_placeholder = ttk.Label(self.tab_visuals, text="Execute o pipeline analítico para projetar os gráficos.", font=('Helvetica', 11, 'italic'), foreground='#64748b')
        self.visuals_placeholder.pack(expand=True)

    def safe_gui_call(self, func, *args, **kwargs) -> None:
        """Agenda de forma segura a execução de tarefas visuais de Tkinter na Thread Principal."""
        if self.root.winfo_exists():
            self.root.after(0, lambda: func(*args, **kwargs))

    def update_log_display(self, message: str):
        """Insere mensagens no painel de terminal de forma assíncrona e thread-safe."""
        def _insert():
            if hasattr(self, 'log_text') and self.log_text.winfo_exists():
                self.log_text.insert(tk.END, message + "\n")
                self.log_text.see(tk.END)
        self.safe_gui_call(_insert)

    def browse_file(self):
        """Abre explorador de arquivos para escolher a planilha."""
        file_selected = filedialog.askopenfilename(
            title="Selecionar Planilha de Entrada",
            filetypes=[("Arquivos de Dados", "*.csv *.xlsx *.xls *.json"), ("Planilhas Excel", "*.xlsx *.xls"), ("Ficheiros CSV", "*.csv"), ("JSON estruturado", "*.json")]
        )
        if file_selected:
            self.filepath = file_selected
            self.file_path_var.set(os.path.basename(file_selected))
            self.status_var.set("PRONTO PARA EXECUÇÃO")
            self.btn_export.configure(state='disabled')

    def load_demo(self):
        """Configura o pipeline para usar os dados sintéticos de demonstração."""
        self.filepath = "__DEMO__"
        self.file_path_var.set("MODO DEMONSTRAÇÃO ATIVO")
        self.status_var.set("DEMO READY")
        self.btn_export.configure(state='disabled')
        self.engine._write_log("Dados sintéticos de teste selecionados.")

    def run_analysis_thread(self):
        """Evita o congelamento da interface gráfica rodando a computação pesada em thread paralela."""
        if not self.filepath:
            messagebox.showwarning("Aviso", "Por favor, carregue um ficheiro ou ative a massa de testes de Demonstração.")
            return
            
        self.btn_run.configure(state='disabled')
        self.status_var.set("PROCESSANDO...")
        
        # Leitura segura dos controlos visualizáveis (BooleanVar) na Thread Principal
        impute_val = self.opt_impute.get()
        outliers_val = self.opt_outliers.get()
        infer_val = self.opt_infer.get()
        
        thread = threading.Thread(
            target=self.execute_pipeline_process,
            args=(impute_val, outliers_val, infer_val)
        )
        thread.daemon = True
        thread.start()

    def execute_pipeline_process(self, impute_nulls: bool, detect_outliers: bool, infer_types: bool):
        """Função executora do pipeline analítico em background (Thread de Trabalho)."""
        try:
            if self.filepath == "__DEMO__":
                self.engine.load_demo_data()
            else:
                self.engine.load_data(self.filepath)
                
            self.engine.run_pipeline(
                impute_nulls=impute_nulls,
                detect_outliers=detect_outliers,
                infer_types=infer_types
            )
            
            # Geração das análises e correlações locais
            self.report_data = self.engine.generate_local_insights()
            
            # Notifica sucesso de forma segura
            self.safe_gui_call(self._on_pipeline_success)
            
        except Exception as e:
            # Notifica erro de forma segura
            self.safe_gui_call(self._on_pipeline_failure, str(e))
        finally:
            self.safe_gui_call(self.btn_run.configure, state='normal')

    def _on_pipeline_success(self):
        """Atualizações de UI de Sucesso executadas estritamente na Thread Principal."""
        self.status_var.set("Pipeline CONCLUÍDO")
        self.btn_export.configure(state='normal')
        
        # Renderiza os gráficos gerados na aba de visualização
        self.render_visuals()
        
        # Força o foco para a aba de Visualização para feedback imediato
        self.notebook.select(1)
        
        messagebox.showinfo("Sucesso", "Análise estatística e gráficos gerados com sucesso!")

    def _on_pipeline_failure(self, error_msg: str):
        """Atualizações de UI de Falha executadas estritamente na Thread Principal."""
        self.status_var.set("FALHA")
        self.engine._write_log(f"Erro fatal: {error_msg}", "ERROR")
        messagebox.showerror("Erro Crítico", f"Falha ao executar pipeline analítico:\n{error_msg}")

    def render_visuals(self):
        """Limpa o placeholder e projeta os gráficos gerados diretamente na aba 'Visualizações'."""
        # Limpar widgets antigos na aba
        for widget in self.tab_visuals.winfo_children():
            widget.destroy()

        self.chart_canvases.clear()

        if not HAS_MATPLOTLIB or not self.engine.figures:
            err_label = ttk.Label(self.tab_visuals, text="Não foi possível projetar visualizações gráficas (Matplotlib indisponível ou dados numéricos nulos).", font=('Helvetica', 10, 'bold'), foreground='#f43f5e')
            err_label.pack(expand=True)
            return

        # Container principal com scroll para os múltiplos gráficos
        canvas_container = tk.Canvas(self.tab_visuals, bg='#0f172a', highlightthickness=0)
        canvas_container.pack(side='left', fill='both', expand=True)

        scrollbar = ttk.Scrollbar(self.tab_visuals, orient='vertical', command=canvas_container.yview)
        scrollbar.pack(side='right', fill='y')

        scroll_frame = tk.Frame(canvas_container, bg='#0f172a')
        
        # Vincular tamanho do frame interno de scroll ao canvas
        scroll_frame.bind(
            "<Configure>",
            lambda e: canvas_container.configure(scrollregion=canvas_container.bbox("all"))
        )
        canvas_container.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas_container.configure(yscrollcommand=scrollbar.set)

        # Injeta cada gráfico gerado sequencialmente de forma responsiva
        for chart_name, fig in self.engine.figures.items():
            card = ttk.Frame(scroll_frame, style='Card.TFrame', padding=10)
            card.pack(fill='x', expand=True, pady=10, padx=10)
            
            # Incorporação direta da figura do Matplotlib no Tkinter
            canvas = FigureCanvasTkAgg(fig, master=card)
            canvas_widget = canvas.get_tk_widget()
            canvas_widget.pack(fill='both', expand=True)
            canvas.draw()
            self.chart_canvases.append(canvas)

    def export_report(self):
        """Abre janela para salvar o relatório consolidado em HTML rico com gráficos."""
        if not self.report_data:
            return
            
        output_path = filedialog.asksaveasfilename(
            title="Salvar Relatório Executivo HTML",
            defaultextension=".html",
            filetypes=[("Relatório Executivo HTML", "*.html")],
            initialfile="relatorio_executivo_aegis.html"
        )
        if output_path:
            try:
                self.engine.export_html_report(output_path, self.report_data)
                messagebox.showinfo("Exportado", f"O relatório técnico executivo HTML foi gerado e guardado em:\n{output_path}")
            except Exception as e:
                messagebox.showerror("Erro ao Exportar", f"Não foi possível salvar o arquivo:\n{str(e)}")


# OPERAÇÃO CLI & INICIALIZAÇÃO DE AMBIENTES
if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Aegis Analytics - CLI de Engenharia Analítica Avançada Local.")
    parser.add_argument("--file", type=str, help="Caminho do arquivo .csv, .json ou .xlsx para ingestão direta sem GUI.")
    parser.add_argument("--demo", action="store_true", help="Dispara a geração de massa sintética para testes rápidos.")
    parser.add_argument("--output", type=str, default="aegis_report.html", help="Caminho para gravação do relatório HTML de saída.")
    parser.add_argument("--cli", action="store_true", help="Força a execução pura em terminal mesmo sem passar parâmetros de arquivos.")
    
    args = parser.parse_args()
    
    # Se nenhum argumento for passado, ou o utilizador não especificar explicitamente o modo CLI, abrir a Interface Gráfica (GUI)
    if not (args.file or args.demo or args.cli):
        root = tk.Tk()
        app = AegisAppGUI(root)
        root.mainloop()
    else:
        # Execução padrão em consola (CLI)
        engine = AegisAnalyticsEngine()
        try:
            if args.demo:
                engine.load_demo_data()
            else:
                engine.load_data(args.file)
                
            # Executar pipeline de normalização e profiling
            engine.run_pipeline()
            
            # Gerar insights qualitativos locais
            report_data = engine.generate_local_insights()
            
            # Consolidar em arquivo final de auditoria HTML
            engine.export_html_report(args.output, report_data)
            
        except Exception as err:
            print(f"\n[FALHA DE EXECUÇÃO] Ocorreu uma interrupção durante a execução do Aegis CLI: {str(err)}")