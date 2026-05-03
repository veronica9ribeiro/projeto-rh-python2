"""
Projeto M3 — Datas, fusos e cálculos de RH
Saint-Gobain | Roadmap Python RH
 
Entrada : colaboradores.xlsx  (gerado pela função criar_planilha_entrada)
Saída   : relatorio_rh.xlsx   (duas abas: Relatório e Alertas)
 
Bibliotecas: pandas, openpyxl, pytz, numpy
"""


import pandas as pd 
import numpy as np 
import pytz
from datatime import date 
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, PatternFill,Border, Side
from openpyxl.utils import get_column_letter


# 1. DADOS DE ENTRADA (simula a planilha real)

FUSO_POR_UF = {
    "SP": "America/Sao_Paulo",
    "RJ": "America/Sao_Paulo", 
    "MG": "America/Sao_Paulo",
    "ES": "America/Sao_Paulo",
    "AM": "America/Manaus",
    "RR": "America/Manaus",
    "MT": "America/Cuiaba",
    "MS": "America/Campo_Grande",
    "AC": "America/Rio_Branco",
    "RO": "America/Porto_Velho",
    "PA": "America/Belem",
    "AP": "America/Belem",
}
  def criar_planilha_entrada(caminho: str):
    """Gera colaboradores.xlsx com dados fictícios para teste."""
    dados = {
        "matricula": ["SG001", "SG002", "SG003", "SG004", "SG005",
                       "SG006", "SG007", "SG008", "SG009", "SG010"],
        "nome":       ["Ana Lima", "Carlos Souza", "Beatriz Neves", "Diego Ramos",
                       "Fernanda Costa", "Gabriel Alves", "Helena Martins",
                       "Igor Santos", "Juliana Pereira", "Klaus Oliveira"],
        "cargo":      ["Analista RH Sr", "Técnico Manutenção", "Estagiária TI",
                       "Supervisor Produção", "Analista Financeiro", "Operador",
                       "Gerente RH", "Analista TI Pl", "Assistente RH", "Diretor Operações"],
        "uf":         ["SP", "AM", "RJ", "SP", "MG", "AC", "SP", "PA", "SP", "AM"],
        "admissao":   ["15/03/2019", "01/08/2021", "10/11/2022", "20/06/2020",
                       "05/01/2018", "12/07/2023", "03/09/2015", "28/02/2022",
                       "17/04/2024", "09/10/2017"],
        "salario":    [6800, 4200, 1600, 5500, 7200, 3100, 12000, 6100, 2800, 18500],
        "tipo":       ["CLT", "CLT", "Estágio", "CLT", "CLT",
                       "CLT", "CLT", "CLT", "CLT", "CLT"],
        "registro_ponto": [
            "2024-05-03 08:28", "2024-05-03 08:30", "2024-05-03 09:05",
            "2024-05-03 07:55", "2024-05-03 08:31", "2024-05-03 08:00",
            "2024-05-03 08:15", "2024-05-03 08:45", "2024-05-03 09:12",
            "2024-05-03 07:48",
        ],
    }
    df = pd.DataFrame(dados)
    df.to_excel(caminho, index=False, sheet_name="Colaboradores")
    print(f"[OK] Planilha de entrada criada: {caminho}")
    return caminho

# 2. LEITURA E LIMPEZA

def ler_colaboradores(caminho: str) -> pd.DataFrame:
    df = pd.read_excel(caminho, sheet_name="Colaboradores", dtype={"matricula": str})

# Converter datas — format explícito evita inversão dia/mês
   df["admissao"] = pd.to_datatime(df["admissao"], format="%d/%m/%Y")
   df["registro_ponto"] = pd.to_datatime(df["registro_ponto"])

 # Normalizar UF

   df["uf"] = df["uf"].str.upper().str.strip()
   df["fuso"] = df["uf"].map(FUSO_POR_UF).fillna("America/Sao_Paulo")

   return df 


# 3. CÁLCULOS DE RH


def calcular_tempo_casa(def: pd.DataFrame) -> pd.DataFrame:
    hoje = pd.Timestamp(date.today())
    delta = hoje - df["admissao"]
    df["dias_empresa"] =  delta.dt.days
    df["anos_empresa"] = df["dias_empresa"] // 365
    df["meses_empresa"] = (df["dias_empresa"] % 365) // 30
    df["tempo_casa"] = (
        df["anos_empresa"].astype(str) + "a" +
        df["meses_empresa"].astype(str) + "m"
    )

    return df



def calcular_aviso_previo(df: pd.DataFrame) -> pd.DataFrame:
    """CLT: 30 dias + 3 por ano de empresa, máximo 90 dias."""
    df["aviso_dias"] = (30 + df["anos_empresa"] * 3).clip(upper=90).astype(int)
    df["data_saida_aviso"] = pd.Timestamp(date.today()) + pd.to_timedelta(
        df["aviso_dias"], unit="D"
    )
    return df
 
 
def calcular_ferias(df: pd.DataFrame) -> pd.DataFrame:
    """Período aquisitivo atual e se há férias vencidas."""
    hoje = pd.Timestamp(date.today())
 
    df["ini_aquisitivo"] = df.apply(
        lambda r: r["admissao"] + pd.DateOffset(years=int(r["anos_empresa"])),
        axis=1,
    )
    df["fim_aquisitivo"] = df["ini_aquisitivo"] + pd.DateOffset(years=1)
 
    # Férias vencidas = período aquisitivo completo mas não gozado
    df["ferias_vencidas"] = df["fim_aquisitivo"] < hoje
    return df
 
 
def calcular_ponto_com_fuso(df: pd.DataFrame) -> pd.DataFrame:
    """Localiza registro de ponto no fuso da UF e converte pra UTC e SP."""
    tz_sp = pytz.timezone("America/Sao_Paulo")
 
    def localizar(row):
        tz = pytz.timezone(row["fuso"])
        return tz.localize(row["registro_ponto"].to_pydatetime())
 
    df["ponto_local"]  = df.apply(localizar, axis=1)
    df["ponto_utc"]    = df["ponto_local"].apply(lambda x: x.astimezone(pytz.utc))
    df["ponto_sede_sp"] = df["ponto_utc"].apply(lambda x: x.astimezone(tz_sp))
 
    # Horário de entrada esperado: 08:30 horário de SP
    referencia = pd.Timestamp("2024-05-03 08:30:00", tz="America/Sao_Paulo")
    df["minutos_atraso"] = (
        df["ponto_sede_sp"]
        .apply(lambda x: pd.Timestamp(x))
        .sub(referencia)
        .dt.total_seconds()
        .div(60)
        .round()
        .astype(int)
    )
    df["status_ponto"] = df["minutos_atraso"].apply(
        lambda m: "Pontual" if m <= 0 else f"Atraso {m}min"
    )
    return df
 
 
def calcular_dias_uteis_mes(df: pd.DataFrame) -> pd.DataFrame:
    """Dias úteis trabalhados desde o início do mês atual."""
    hoje = date.today()
    inicio_mes = date(hoje.year, hoje.month, 1)
    df["du_no_mes"] = np.busday_count(inicio_mes, hoje)
    return df
 
 

# 4. MONTAR O DATAFRAME FINAL

 
def processar(df: pd.DataFrame) -> pd.DataFrame:
    df = calcular_tempo_casa(df)
    df = calcular_aviso_previo(df)
    df = calcular_ferias(df)
    df = calcular_ponto_com_fuso(df)
    df = calcular_dias_uteis_mes(df)
    return df
 
 

# 5. EXPORTAR RELATÓRIO .XLSX FORMATADO

COLS_RELATORIO = {
    "matricula":        "Matrícula",
    "nome":             "Nome",
    "cargo":            "Cargo",
    "uf":               "UF",
    "tipo":             "Tipo",
    "admissao":         "Admissão",
    "tempo_casa":       "Tempo de Casa",
    "anos_empresa":     "Anos",
    "aviso_dias":       "Aviso Prévio (dias)",
    "data_saida_aviso": "Saída c/ Aviso",
    "ini_aquisitivo":   "Início Aquisitivo",
    "fim_aquisitivo":   "Fim Aquisitivo",
    "ferias_vencidas":  "Férias Vencidas",
    "status_ponto":     "Status Ponto 03/05",
    "du_no_mes":        "DU no Mês",
}
 
COLS_ALERTAS = {
    "matricula":     "Matrícula",
    "nome":          "Nome",
    "cargo":         "Cargo",
    "ferias_vencidas": "Férias Vencidas",
    "status_ponto":  "Status Ponto",
    "fim_aquisitivo": "Vencimento Férias",
}
 
 
def _estilo_cabecalho(ws, row: int, n_cols: int, cor_hex: str):
    fill = PatternFill("solid", start_color=cor_hex, end_color=cor_hex)
    font = Font(bold=True, color="FFFFFF", size=10)
    for col in range(1, n_cols + 1):
        cell = ws.cell(row=row, column=col)
        cell.fill = fill
        cell.font = font
        cell.alignment = Alignment(horizontal="center", vertical="center")
 
 
def _borda_fina():
    lado = Side(style="thin", color="D0D0D0")
    return Border(left=lado, right=lado, top=lado, bottom=lado)
 
 
def _auto_width(ws):
    for col in ws.columns:
        max_len = max((len(str(c.value)) if c.value else 0) for c in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = min(max_len + 4, 36)
 
 
def exportar_xlsx(df: pd.DataFrame, caminho: str):
    hoje_str = date.today().strftime("%d/%m/%Y")
 
    # ── Aba 1: Relatório completo ──
    rel = df[list(COLS_RELATORIO.keys())].copy()
    rel.rename(columns=COLS_RELATORIO, inplace=True)
    for col in ["Admissão", "Saída c/ Aviso", "Início Aquisitivo", "Fim Aquisitivo"]:
        if col in rel.columns:
            rel[col] = pd.to_datetime(rel[col]).dt.strftime("%d/%m/%Y")
    rel["Férias Vencidas"] = rel["Férias Vencidas"].map({True: "SIM", False: "Não"})
 
    # ── Aba 2: Alertas ──
    alertas = df[
        df["ferias_vencidas"] | (df["minutos_atraso"] > 0)
    ][list(COLS_ALERTAS.keys())].copy()
    alertas.rename(columns=COLS_ALERTAS, inplace=True)
    alertas["Férias Vencidas"] = alertas["Férias Vencidas"].map({True: "SIM", False: "Não"})
    alertas["Vencimento Férias"] = pd.to_datetime(alertas["Vencimento Férias"]).dt.strftime("%d/%m/%Y")
 
    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        rel.to_excel(writer, sheet_name="Relatório", index=False, startrow=2)
        alertas.to_excel(writer, sheet_name="Alertas", index=False, startrow=2)
 
    # ── Pós-formatação com openpyxl ──
    wb = load_workbook(caminho)
 
    # — Aba Relatório —
    ws = wb["Relatório"]
    ws["A1"] = f"Relatório RH — Saint-Gobain    |    Gerado em {hoje_str}"
    ws["A1"].font = Font(bold=True, size=12, color="1F4E79")
    ws.merge_cells(f"A1:{get_column_letter(len(COLS_RELATORIO))}1")
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22
    ws.row_dimensions[3].height = 18
 
    _estilo_cabecalho(ws, row=3, n_cols=len(COLS_RELATORIO), cor_hex="2E75B6")
 
    borda = _borda_fina()
    fill_alt = PatternFill("solid", start_color="EBF3FB", end_color="EBF3FB")
    fill_alerta = PatternFill("solid", start_color="FFF2CC", end_color="FFF2CC")
    font_alerta = Font(bold=True, color="7F6000")
 
    idx_ferias = list(COLS_RELATORIO.keys()).index("ferias_vencidas") + 1
    idx_ponto  = list(COLS_RELATORIO.keys()).index("status_ponto") + 1
 
    for i, row in enumerate(ws.iter_rows(min_row=4, max_row=ws.max_row), start=0):
        for cell in row:
            cell.border = borda
            cell.alignment = Alignment(vertical="center")
            if i % 2 == 1:
                cell.fill = fill_alt
        # Destacar férias vencidas
        cell_ferias = ws.cell(row=i + 4, column=idx_ferias)
        if cell_ferias.value == "SIM":
            cell_ferias.fill = fill_alerta
            cell_ferias.font = font_alerta
        # Destacar atrasos
        cell_ponto = ws.cell(row=i + 4, column=idx_ponto)
        if cell_ponto.value and cell_ponto.value != "Pontual":
            cell_ponto.fill = fill_alerta
            cell_ponto.font = font_alerta
 
    ws.freeze_panes = "A4"
    _auto_width(ws)
 
    # — Aba Alertas —
    ws2 = wb["Alertas"]
    ws2["A1"] = f"Alertas RH — colaboradores que precisam de atenção    |    {hoje_str}"
    ws2["A1"].font = Font(bold=True, size=12, color="833C00")
    ws2.merge_cells(f"A1:{get_column_letter(len(COLS_ALERTAS))}1")
    ws2["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws2.row_dimensions[1].height = 22
    ws2.row_dimensions[3].height = 18
 
    _estilo_cabecalho(ws2, row=3, n_cols=len(COLS_ALERTAS), cor_hex="C55A11")
 
    for i, row in enumerate(ws2.iter_rows(min_row=4, max_row=ws2.max_row), start=0):
        for cell in row:
            cell.border = borda
            cell.alignment = Alignment(vertical="center")
            if i % 2 == 1:
                cell.fill = PatternFill("solid", start_color="FCE4D6", end_color="FCE4D6")
 
    ws2.freeze_panes = "A4"
    _auto_width(ws2)
 
    wb.save(caminho)
    print(f"[OK] Relatório exportado: {caminho}")
    print(f"     Colaboradores processados : {len(df)}")
    print(f"     Alertas gerados           : {len(alertas)}")
 
 

# 6. PIPELINE PRINCIPAL

 
if __name__ == "__main__":
    ENTRADA  = "/home/claude/colaboradores.xlsx"
    SAIDA    = "/home/claude/relatorio_rh.xlsx"
 
    criar_planilha_entrada(ENTRADA)
    df = ler_colaboradores(ENTRADA)
    df = processar(df)
    exportar_xlsx(df, SAIDA)





  
    
