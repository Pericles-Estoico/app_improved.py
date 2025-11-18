# app_improved.py
import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO, StringIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import requests

# ==============================================================================
# CONFIGURAÇÃO GERAL
# ==============================================================================

st.set_page_config(
    page_title="Pure & Posh Baby - Relatórios & Produção",
    page_icon="👑",
    layout="wide"
)

st.markdown(
    """
    <style>
    .centered-title { text-align: center; width: 100%; margin: 0 auto; }
    @media (max-width: 768px) { .centered-title { text-align: center; } }
    .subtitle { color: #666; font-size: 0.9rem; }
    .box { padding: 0.8rem 1rem; border-radius: 8px; margin-bottom: 0.8rem; }
    .box-info { background: #e8f4ff; border-left: 4px solid #1e88e5; }
    .box-warning { background: #fff8e1; border-left: 4px solid #ffb300; }
    .box-success { background: #e8f5e9; border-left: 4px solid #43a047; }
    .section-title { font-size: 1.2rem; font-weight: 700; margin-top: 1rem; }
    </style>
    """,
    unsafe_allow_html=True
)

st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios & Planejamento de Produção")
st.markdown("**Pure & Posh Baby — Vendas → Estoque → Produção**")
st.markdown('</div>', unsafe_allow_html=True)

# ==============================================================================
# CONSTANTES
# ==============================================================================

# 🔗 Planilha Mãe: você carrega por upload (xlsx)
# 🔗 template_estoque: leitura apenas, via link público da planilha já usada no cockpit
TEMPLATE_ESTOQUE_URL = (
    "https://docs.google.com/spreadsheets/"
    "d/1PpiMQingHf4llA03BiPIuPJPIZqul4grRU_emWDEK1o/export?format=csv&gid=1456159896"
)

# ==============================================================================
# ESTADO DA SESSÃO
# ==============================================================================

if "planilha_mae_carregada" not in st.session_state:
    st.session_state["planilha_mae_carregada"] = False
if "df_mae" not in st.session_state:
    st.session_state["df_mae"] = None

# ==============================================================================
# FUNÇÕES AUXILIARES
# ==============================================================================

def safe_int(x, default=0):
    """Converte qualquer coisa para int sem quebrar."""
    try:
        if x is None or (isinstance(x, float) and np.isnan(x)):
            return default
        s = str(x).strip()
        if s == "" or s.lower() in {"nan", "none", "null"}:
            return default
        return int(float(s.replace(",", ".")))
    except Exception:
        return default


def estoque_ajustado(v):
    """
    ⚠️ Regra importante:
    - Se estoque tiver negativo, trata como 0 para cálculo de FALTA.
    - Isso vale tanto para produto pronto quanto para semi/gola/bordado.
    """
    return max(safe_int(v, 0), 0)


@st.cache_data
def load_excel(arquivo):
    """Carrega um arquivo Excel em um DataFrame, com cache para performance."""
    return pd.read_excel(arquivo)


@st.cache_data(ttl=60)
def carregar_template_estoque():
    """
    Lê a aba 'template_estoque' (gid fixo) só para CONSULTA.
    Não altera nada no Google Sheets.
    """
    try:
        r = requests.get(TEMPLATE_ESTOQUE_URL, timeout=20)
        r.raise_for_status()
        df = pd.read_csv(StringIO(r.text))
        df.columns = df.columns.str.strip().str.lower()
        # garantias mínimas
        if "codigo" not in df.columns:
            df["codigo"] = ""
        if "nome" not in df.columns:
            df["nome"] = ""
        if "estoque_atual" not in df.columns:
            df["estoque_atual"] = 0
        return df
    except Exception as e:
        st.error(f"Erro ao carregar template_estoque: {e}")
        return pd.DataFrame(columns=["codigo", "nome", "estoque_atual"])


def get_categoria_ordem(semi):
    """
    Define ordem de exibição:
    1 = Manga Longa
    2 = Manga Curta Menina
    3 = Manga Curta Menino
    4 = Mijão
    5 = Outros
    + cor + tamanho
    """
    s = str(semi).lower()

    # Categoria
    if "manga longa" in s:
        categoria = 1
    elif "manga curta" in s and "menina" in s:
        categoria = 2
    elif "manga curta" in s and "menino" in s:
        categoria = 3
    elif "mijão" in s or "mijao" in s:
        categoria = 4
    else:
        categoria = 5

    # Cor
    if "branco" in s:
        cor = 1
    elif "off-white" in s or "off white" in s:
        cor = 2
    elif "rosa" in s:
        cor = 3
    elif "azul" in s:
        cor = 4
    elif "vermelho" in s:
        cor = 5
    elif "marinho" in s:
        cor = 6
    else:
        cor = 7

    # Tamanho
    if "-rn" in s or " rn" in s:
        tam = 1
    elif "-p" in s or " p" in s:
        tam = 2
    elif "-m" in s or " m" in s:
        tam = 3
    elif "-g" in s or " g" in s:
        tam = 4
    else:
        tam = 5

    return categoria, cor, tam


def explodir_kits(df_vendas_com_mae, df_mae_completa):
    """
    Mesma lógica do app antigo:
    - Recebe vendas já mescladas com a planilha mãe.
    - Usa coluna 'componentes_codigos' para decompor kits em produtos filhos.
    - Retorna DataFrame com colunas: semi, gola, bordado, quantidade.
    """
    componentes_finais = []

    df_mae_completa = df_mae_completa.set_index("codigo")

    def obter_componentes(codigo, quantidade):
        lista_componentes_recursiva = []
        try:
            produto = df_mae_completa.loc[codigo]
        except KeyError:
            return []

        # 1) componente direto (semi/gola/bordado)
        semi_valido = False
        if "semi" in produto.index:
            if pd.notna(produto["semi"]):
                if isinstance(produto["semi"], str) and produto["semi"].strip() != "":
                    semi_valido = True

        if semi_valido:
            lista_componentes_recursiva.append(
                {
                    "semi": produto["semi"],
                    "gola": produto["gola"] if pd.notna(produto["gola"]) else "",
                    "bordado": produto["bordado"] if pd.notna(produto["bordado"]) else "",
                    "quantidade": quantidade,
                }
            )

        # 2) componentes aninhados (kits dentro de kits)
        componentes_codigos_valido = False
        if "componentes_codigos" in produto.index:
            if pd.notna(produto["componentes_codigos"]):
                componentes_str = str(produto["componentes_codigos"]).strip()
                if componentes_str != "" and componentes_str.lower() != "nan":
                    componentes_codigos_valido = True

        if componentes_codigos_valido:
            codigos_aninhados = str(produto["componentes_codigos"]).split(";")
            for cod_aninhado in codigos_aninhados:
                cod_aninhado = cod_aninhado.strip()
                if cod_aninhado:
                    lista_componentes_recursiva.extend(
                        obter_componentes(cod_aninhado, quantidade)
                    )

        return lista_componentes_recursiva

    for _, venda in df_vendas_com_mae.iterrows():
        componentes_finais.extend(
            obter_componentes(venda["codigo"], safe_int(venda["quantidade"], 0))
        )

    return pd.DataFrame(componentes_finais)


def gerar_excel_hierarquico_semis_golas(df_semis, df_golas):
    """
    Gera um Excel hierárquico:
    - Linha de Semi
    - Embaixo, as golas correspondentes.
    Com colunas: Item, Qtd Necessária, Estoque Atual, Falta
    """
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Produzir Hoje"

    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    headers = ["Item", "Qtd Necessária", "Estoque Atual", "Falta"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

    # monta estrutura hierárquica em memória
    linhas = []
    for _, semi_row in df_semis.iterrows():
        semi_nome = semi_row["Semi"]
        linhas.append(
            {
                "Item": semi_nome,
                "Qtd Necessária": semi_row["Qtd Necessária"],
                "Estoque Atual": semi_row["Estoque Atual"],
                "Falta": semi_row["Falta"],
                "is_semi": True,
            }
        )

        sub = df_golas[df_golas["Semi"] == semi_nome]
        for _, gola_row in sub.iterrows():
            linhas.append(
                {
                    "Item": f"  Gola: {gola_row['Gola']}",
                    "Qtd Necessária": gola_row["Qtd Necessária"],
                    "Estoque Atual": gola_row["Estoque Atual"],
                    "Falta": gola_row["Falta"],
                    "is_semi": False,
                }
            )

    # escreve no Excel
    row_num = 2
    for linha in linhas:
        for col_num, key in enumerate(headers, 1):
            cell = ws.cell(row=row_num, column=col_num, value=linha[key])
            cell.border = border
        row_num += 1

    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 10

    wb.save(output)
    output.seek(0)
    return output


def gerar_excel_simples(df, sheet_name="Relatorio"):
    """
    Gera Excel simples, tabela direta (sem hierarquia).
    """
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name

    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    headers = list(df.columns)
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.border = border

    for r_idx, row in enumerate(df.itertuples(index=False), 2):
        for c_idx, value in enumerate(row, 1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            cell.border = border

    for column in ws.columns:
        max_length = 0
        col_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception:
                pass
        ws.column_dimensions[col_letter].width = min(max_length + 2, 60)

    wb.save(output)
    output.seek(0)
    return output

# ==============================================================================
# 1. CONFIGURAÇÃO INICIAL — PLANILHA MÃE
# ==============================================================================

st.markdown('<div class="section-title">1. Configuração Inicial — Planilha Mãe</div>', unsafe_allow_html=True)
st.markdown(
    """
    <div class="box box-info">
    <b>O que é a Planilha Mãe?</b><br>
    • <code>codigo</code> → código vendido (SKU do produto pronto)<br>
    • <code>semi</code> → Semi usado para produzir o body<br>
    • <code>gola</code> → Gola usada (quando existir)<br>
    • <code>bordado</code> → Bordado da gola (quando a gola não existe pronta)<br>
    • <code>componentes_codigos</code> → códigos filhos que compõem kits (opcional)<br><br>
    Você carrega uma vez por sessão. Pode ir completando com o tempo.
    </div>
    """,
    unsafe_allow_html=True,
)


def carregar_planilha_mae(arquivo):
    try:
        with st.spinner("Carregando e validando Planilha Mãe..."):
            df = load_excel(arquivo)
            df.columns = (
                df.columns.str.strip()
                .str.replace(" ", "_")
                .str.lower()
            )

            colunas_essenciais = ["codigo", "semi", "gola", "bordado"]
            if not all(col in df.columns for col in colunas_essenciais):
                st.error(
                    "❌ A Planilha Mãe deve conter as colunas: "
                    + ", ".join(colunas_essenciais)
                )
                return

            if "componentes_codigos" not in df.columns:
                df["componentes_codigos"] = ""

            st.session_state["df_mae"] = df
            st.session_state["planilha_mae_carregada"] = True
            st.success(f"✅ Planilha Mãe carregada: {len(df)} produtos cadastrados.")
            st.rerun()
    except Exception as e:
        st.error(f"Erro ao carregar planilha mãe: {e}")


if st.session_state["planilha_mae_carregada"]:
    st.success(
        f"✅ Planilha Mãe carregada: {len(st.session_state['df_mae'])} produtos cadastrados."
    )
    with st.expander("🔄 Recarregar / Atualizar Planilha Mãe"):
        uploaded_mae_nova = st.file_uploader(
            "Substituir Planilha Mãe", type=["xlsx"], key="planilha_mae_nova"
        )
        if uploaded_mae_nova:
            carregar_planilha_mae(uploaded_mae_nova)
else:
    st.info(
        "📋 Para começar, carregue a Planilha Mãe (`codigo`, `semi`, `gola`, "
        "`bordado`, `componentes_codigos`)."
    )
    uploaded_mae = st.file_uploader(
        "Carregar Planilha Mãe", type=["xlsx"], key="planilha_mae"
    )
    if uploaded_mae:
        carregar_planilha_mae(uploaded_mae)

# Se não tiver planilha mãe, para aqui.
if not st.session_state["planilha_mae_carregada"]:
    st.stop()

df_mae = st.session_state["df_mae"]

# ==============================================================================
# 2. PROCESSAR VENDAS DO DIA
# ==============================================================================

st.markdown('<div class="section-title">2. Processar Vendas do Dia</div>', unsafe_allow_html=True)
st.markdown(
    """
    <div class="box box-info">
    <b>Formato da planilha de vendas:</b><br>
    • Excel (<code>.xlsx</code>)<br>
    • Colunas obrigatórias: <b>Código</b> e <b>Quantidade</b><br><br>
    <b>Fluxo:</b><br>
    1) Soma vendas por código;<br>
    2) Consulta estoque de produto pronto na <code>template_estoque</code>;<br>
    3) Usa o que já tem pronto;<br>
    4) Só explode em insumos o que realmente falta produzir.<br><br>
    <b>Produzir Hoje</b> já considera: se o estoque estiver negativo, ele é tratado como 0.
    Ou seja, o relatório mostra apenas o que o <b>dia de vendas</b> exige.
    </div>
    """,
    unsafe_allow_html=True,
)

uploaded_vendas = st.file_uploader(
    "📈 Planilha de Vendas (Mercado Livre / Dia)", type=["xlsx"], key="vendas"
)

if not uploaded_vendas:
    st.stop()

# ---------------- VENDAS -----------------
try:
    with st.spinner("Processando vendas..."):
        df_vendas = load_excel(uploaded_vendas)
        df_vendas.columns = (
            df_vendas.columns.str.strip()
            .str.replace(" ", "_")
            .str.lower()
        )

        if "código" in df_vendas.columns:
            df_vendas = df_vendas.rename(columns={"código": "codigo"})

        if "codigo" not in df_vendas.columns or "quantidade" not in df_vendas.columns:
            st.error("❌ A planilha de vendas deve ter as colunas 'Código' e 'Quantidade'.")
            st.stop()

        df_vendas["codigo"] = df_vendas["codigo"].astype(str).str.strip()
        df_vendas["quantidade"] = df_vendas["quantidade"].apply(safe_int)

        df_vendas = (
            df_vendas.groupby("codigo", as_index=False)["quantidade"].sum()
        )

except Exception as e:
    st.error(f"Erro ao ler planilha de vendas: {e}")
    st.stop()

# ---------------- ESTOQUE PRONTO -----------------
df_estoque = carregar_template_estoque()

if df_estoque.empty:
    st.error("Não foi possível ler a template_estoque para consultar estoque.")
    st.stop()

# mapa de estoque por código de PRODUTO PRONTO
df_estoque_produtos = df_estoque.copy()
df_estoque_produtos["codigo"] = df_estoque_produtos["codigo"].astype(str).str.strip()
estoque_prod_map = {
    row["codigo"]: estoque_ajustado(row["estoque_atual"])
    for _, row in df_estoque_produtos.iterrows()
}

# tabela de resumo por produto
linhas_resumo = []
linhas_para_produzir = []

for _, row in df_vendas.iterrows():
    cod = str(row["codigo"]).strip()
    qtd = safe_int(row["quantidade"], 0)
    est = estoque_prod_map.get(cod, 0)      # já vem ajustado (negativo->0)
    usa_pronto = min(qtd, est)
    falta_produto = max(qtd - usa_pronto, 0)

    linhas_resumo.append(
        {
            "codigo": cod,
            "Vendido": qtd,
            "Estoque Produto Pronto": est,
            "Atendido com Pronto": usa_pronto,
            "Para Produzir": falta_produto,
        }
    )

    if falta_produto > 0:
        linhas_para_produzir.append({"codigo": cod, "quantidade": falta_produto})

df_resumo_produtos = pd.DataFrame(linhas_resumo)
df_vendas_para_produzir = pd.DataFrame(linhas_para_produzir)

st.subheader("2.1 Situação dos Produtos Prontos (template_estoque)")

st.markdown(
    """
    <div class="subtitle">
    • <b>Vendido</b>: total da planilha de vendas.<br>
    • <b>Estoque Produto Pronto</b>: quantidade atual lida da <code>template_estoque</code>,
      já tratando negativos como 0.<br>
    • <b>Atendido com Pronto</b>: quanto dessa venda é coberta com estoque pronto.<br>
    • <b>Para Produzir</b>: só o que não tem pronto e precisa ir para produção.
    </div>
    """,
    unsafe_allow_html=True,
)

st.dataframe(df_resumo_produtos, use_container_width=True, height=320)

st.download_button(
    "📥 Baixar resumo de Produtos Prontos (CSV)",
    df_resumo_produtos.to_csv(index=False, encoding="utf-8-sig"),
    file_name="resumo_produtos_prontos.csv",
    mime="text/csv",
)

if df_vendas_para_produzir.empty:
    st.success("🎉 Todas as vendas do dia foram atendidas com produto pronto. Nada para produzir hoje.")
    st.stop()

# ==============================================================================
# 3. EXPLOSÃO EM INSUMOS (SEMI / GOLA / BORDADO)
# ==============================================================================

st.subheader("2.2 Explosão em Insumos (apenas o que falta produzir)")

# Mescla com planilha mãe
df_merged = pd.merge(
    df_vendas_para_produzir,
    df_mae,
    on="codigo",
    how="left",
    suffixes=("", "_mae"),
)

codigos_faltando_na_mae = df_merged[df_merged["semi"].isna()]["codigo"].unique()
dados_validos_df = df_merged.dropna(subset=["semi"]).copy()

if len(codigos_faltando_na_mae) > 0:
    st.warning(
        f"⚠️ {len(codigos_faltando_na_mae)} código(s) não encontrados na Planilha Mãe. "
        "Eles não entram na explosão até serem cadastrados."
    )
    df_falt_mae = pd.DataFrame({"codigo": codigos_faltando_na_mae})
    st.dataframe(df_falt_mae, use_container_width=True, height=180)
else:
    st.info("✅ Todos os códigos enviados existem na Planilha Mãe.")

if dados_validos_df.empty:
    st.error("Nenhum código com semi cadastrado para explodir.")
    st.stop()

with st.spinner("Explodindo kits e montando insumos..."):
    df_componentes = explodir_kits(dados_validos_df, df_mae)

st.success(
    f"✅ Explosão concluída: {len(df_componentes)} linhas de componentes (semi/gola/bordado)."
)

# ==============================================================================
# 4. CRUZAR INSUMOS x ESTOQUE (template_estoque)
# ==============================================================================

st.markdown('<div class="section-title">3. Produzir Hoje — Planejamento por Insumo</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="box box-warning">
    <b>Regra de Estoque (importantíssima):</b><br>
    • Se o <code>estoque_atual</code> da template_estoque estiver <b>negativo</b>,
      ele é tratado como <b>0</b> para o cálculo de <b>Falta</b>.<br>
    • Ou seja, <b>Produzir Hoje</b> mostra o que o <u>dia de vendas</u> exige.<br>
    • O rombo antigo (negativo histórico) continua sendo tratado no cockpit de estoque.
    </div>
    """,
    unsafe_allow_html=True,
)

# mapa de estoque por NOME (insumos)
df_estoque_ins = df_estoque.copy()
df_estoque_ins["nome_norm"] = df_estoque_ins["nome"].astype(str).str.strip().str.lower()
estoque_ins_map = {
    row["nome_norm"]: estoque_ajustado(row["estoque_atual"])
    for _, row in df_estoque_ins.iterrows()
}

# normalizar nomes nos componentes
df_comp = df_componentes.copy()
df_comp["semi"] = df_comp["semi"].fillna("").astype(str)
df_comp["gola"] = df_comp["gola"].fillna("").astype(str)
df_comp["bordado"] = df_comp["bordado"].fillna("").astype(str)

df_comp["semi_norm"] = df_comp["semi"].str.strip().str.lower()
df_comp["gola_norm"] = df_comp["gola"].str.strip().str.lower()
df_comp["bordado_norm"] = df_comp["bordado"].str.strip().str.lower()

# ---------------- SEMIS -----------------
semis = df_comp[df_comp["semi"] != ""].copy()
if semis.empty:
    st.error("Nenhum semi encontrado na explosão.")
    st.stop()

df_semis = (
    semis.groupby("semi_norm", as_index=False)["quantidade"].sum()
    .rename(columns={"quantidade": "Qtd Necessária"})
)

# pegar nome original e ordem
nome_map_semi = (
    semis.groupby("semi_norm")["semi"].first().to_dict()
)
df_semis["Semi"] = df_semis["semi_norm"].map(nome_map_semi)

# estoque atual ajustado (negativo → 0)
df_semis["Estoque Atual"] = df_semis["semi_norm"].map(
    lambda n: estoque_ins_map.get(n, 0)
)
df_semis["Estoque Atual"] = df_semis["Estoque Atual"].apply(estoque_ajustado)

df_semis["Falta"] = (df_semis["Qtd Necessária"] - df_semis["Estoque Atual"]).clip(lower=0)

# ordem
ordens = df_semis["Semi"].apply(lambda x: pd.Series(get_categoria_ordem(x)))
df_semis[["cat", "cor", "tam"]] = ordens
df_semis = df_semis.sort_values(["cat", "cor", "tam", "Semi"]).reset_index(drop=True)

# manter só colunas finais
df_semis_view = df_semis[["Semi", "Qtd Necessária", "Estoque Atual", "Falta"]]

st.markdown("### 3.1 Produzir Hoje – **SEMIS**")
st.dataframe(df_semis_view, use_container_width=True, height=320)

# ---------------- GOLAS -----------------
golas = df_comp[(df_comp["gola"] != "")].copy()

if not golas.empty:
    df_golas = (
        golas.groupby(["semi_norm", "gola_norm"], as_index=False)["quantidade"].sum()
        .rename(columns={"quantidade": "Qtd Necessária"})
    )

    nome_map_gola = golas.groupby("gola_norm")["gola"].first().to_dict()
    df_golas["Semi"] = df_golas["semi_norm"].map(nome_map_semi)
    df_golas["Gola"] = df_golas["gola_norm"].map(nome_map_gola)

    df_golas["Estoque Atual"] = df_golas["gola_norm"].map(
        lambda n: estoque_ins_map.get(n, 0)
    )
    df_golas["Estoque Atual"] = df_golas["Estoque Atual"].apply(estoque_ajustado)
    df_golas["Falta"] = (
        df_golas["Qtd Necessária"] - df_golas["Estoque Atual"]
    ).clip(lower=0)

    ordens_g = df_golas["Semi"].apply(lambda x: pd.Series(get_categoria_ordem(x)))
    df_golas[["cat", "cor", "tam"]] = ordens_g
    df_golas = df_golas.sort_values(
        ["cat", "cor", "tam", "Semi", "Gola"]
    ).reset_index(drop=True)

    df_golas_view = df_golas[["Semi", "Gola", "Qtd Necessária", "Estoque Atual", "Falta"]]

    st.markdown("### 3.2 Produzir Hoje – **GOLAS** (casadas com os SEMIS)")
    st.dataframe(df_golas_view, use_container_width=True, height=320)
else:
    df_golas = pd.DataFrame(columns=["Semi", "Gola", "Qtd Necessária", "Estoque Atual", "Falta"])
    df_golas_view = df_golas.copy()
    st.info("Nenhuma gola encontrada na explosão.")

# ---------------- BORDADOS -----------------
# Regra: só entra aqui quando NÃO tem gola e tem bordado
bordados = df_comp[
    (df_comp["gola_norm"] == "") & (df_comp["bordado_norm"] != "")
].copy()

if not bordados.empty:
    df_bord = (
        bordados.groupby("bordado_norm", as_index=False)["quantidade"].sum()
        .rename(columns={"quantidade": "Qtd Necessária"})
    )
    nome_map_bord = (
        bordados.groupby("bordado_norm")["bordado"].first().to_dict()
    )
    df_bord["Bordado"] = df_bord["bordado_norm"].map(nome_map_bord)

    df_bord["Estoque Atual"] = df_bord["bordado_norm"].map(
        lambda n: estoque_ins_map.get(n, 0)
    )
    df_bord["Estoque Atual"] = df_bord["Estoque Atual"].apply(estoque_ajustado)
    df_bord["Falta"] = (
        df_bord["Qtd Necessária"] - df_bord["Estoque Atual"]
    ).clip(lower=0)

    df_bord_view = df_bord[["Bordado", "Qtd Necessária", "Estoque Atual", "Falta"]]

    st.markdown("### 3.3 Produzir Hoje – **BORDADOS** (quando não há gola pronta)")
    st.dataframe(df_bord_view, use_container_width=True, height=260)
else:
    df_bord_view = pd.DataFrame(columns=["Bordado", "Qtd Necessária", "Estoque Atual", "Falta"])
    st.info("Nenhum bordado usado diretamente (sem gola) neste dia.")

# ==============================================================================
# 5. DOWNLOADS
# ==============================================================================

st.markdown("### 4. Relatórios para Download")

col_a, col_b, col_c = st.columns(3)

with col_a:
    # Hierárquico: Semi + golas logo abaixo
    excel_semis_golas = gerar_excel_hierarquico_semis_golas(
        df_semis_view, df_golas_view
    )
    st.download_button(
        "📥 Baixar 'Produzir Hoje – Semis + Golas'",
        data=excel_semis_golas,
        file_name="produzir_hoje_semis_golas.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col_b:
    excel_semis = gerar_excel_simples(df_semis_view, sheet_name="Semis")
    st.download_button(
        "📥 Baixar 'Produzir Hoje – Semis' (tabela simples)",
        data=excel_semis,
        file_name="produzir_hoje_semis_simples.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

with col_c:
    excel_bord = gerar_excel_simples(df_bord_view, sheet_name="Bordados")
    st.download_button(
        "📥 Baixar 'Produzir Hoje – Bordados'",
        data=excel_bord,
        file_name="produzir_hoje_bordados.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

st.markdown(
    """
    <div class="box box-success">
    <b>Resumo do que cada relatório faz:</b><br>
    • <b>Produzir Hoje – Semis + Golas</b>: visão hierárquica, cada Semi com suas golas logo abaixo (ideal para produção).<br>
    • <b>Produzir Hoje – Semis (simples)</b>: apenas tabela de Semis, para quem quer algo direto.<br>
    • <b>Produzir Hoje – Bordados</b>: lista dos bordados necessários quando não há gola pronta cadastrada.<br><br>
    Tudo já calculado considerando que estoque negativo é tratado como 0 para a conta de <b>Falta</b>.
    </div>
    """,
    unsafe_allow_html=True,
)
