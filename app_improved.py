import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import requests

# ==============================================================================
# CONFIGURAÇÃO GERAL DO APP
# ==============================================================================

st.set_page_config(
    page_title="Pure & Posh Baby - Relatórios & Produção",
    page_icon="👑",
    layout="wide"
)

# ------------------------------------------------------------------------------
# CSS básico
# ------------------------------------------------------------------------------
st.markdown(
    """
    <style>
    .centered-title { text-align: center; width: 100%; margin: 0 auto; }
    .explicacao-box {
        background-color: #f8f9fa;
        border-left: 4px solid #0d6efd;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
        font-size: 0.9rem;
    }
    .alerta-box {
        background-color: #fff3cd;
        border-left: 4px solid #ffc107;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
        font-size: 0.9rem;
    }
    .sucesso-box {
        background-color: #d4edda;
        border-left: 4px solid #28a745;
        padding: 0.8rem 1rem;
        border-radius: 6px;
        margin-bottom: 0.8rem;
        font-size: 0.9rem;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# ------------------------------------------------------------------------------
# Header
# ------------------------------------------------------------------------------
st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios & Planejamento de Produção")
st.markdown("**Pure & Posh Baby** — Vendas → Estoque → Produção")
st.markdown('</div>', unsafe_allow_html=True)


# ==============================================================================
# SESSION STATE
# ==============================================================================
if "planilha_mae_carregada" not in st.session_state:
    st.session_state["planilha_mae_carregada"] = False
if "df_mae" not in st.session_state:
    st.session_state["df_mae"] = None


# ==============================================================================
# CONFIG: TEMPLATE_ESTOQUE (Google Sheets - somente leitura)
# ==============================================================================

# ❗ Ajuste este URL se a planilha mudar.
# É o mesmo ID do cockpit (template_estoque).
TEMPLATE_ESTOQUE_CSV_URL = (
    "https://docs.google.com/spreadsheets/d/1PpiMQingHf4llA03BiPIuPJPIZqul4grRU_emWDEK1o/"
    "export?format=csv"
)


# ==============================================================================
# FUNÇÕES CORE
# ==============================================================================

@st.cache_data
def load_excel(arquivo):
    """Carrega um arquivo Excel em um DataFrame, com cache."""
    return pd.read_excel(arquivo)


@st.cache_data(ttl=60)
def carregar_template_estoque():
    """
    Lê a planilha template_estoque em modo SOMENTE LEITURA.
    Espera colunas:
      - codigo
      - nome
      - categoria (ex: Produto, Semi, Gola, Bordado)
      - estoque_atual
    Se tiver mais colunas, elas são ignoradas.
    """
    try:
        r = requests.get(TEMPLATE_ESTOQUE_CSV_URL, timeout=20)
        r.raise_for_status()
        df = pd.read_csv(BytesIO(r.content), encoding="utf-8")

        df.columns = df.columns.str.strip().str.lower()
        # Garante colunas mínimas
        if "codigo" not in df.columns:
            df["codigo"] = ""
        if "nome" not in df.columns:
            df["nome"] = ""
        if "categoria" not in df.columns:
            df["categoria"] = "Produto"
        if "estoque_atual" not in df.columns:
            df["estoque_atual"] = 0

        df["estoque_atual"] = pd.to_numeric(df["estoque_atual"], errors="coerce").fillna(0)

        # Normalizações para matching por nome
        df["nome_norm"] = df["nome"].astype(str).str.strip().str.lower()
        df["categoria_norm"] = df["categoria"].astype(str).str.strip().str.lower()

        return df
    except Exception as e:
        st.warning(f"⚠️ Não foi possível ler a template_estoque: {e}")
        return pd.DataFrame()


def get_categoria_ordem(semi):
    """
    Determina:
      - categoria (1 a 4, para ordenação)
      - cor_ordem
      - tamanho_ordem
    com base no texto do 'semi'.
    ORDEM:
        1) Manga Longa
        2) Manga Curta Menina
        3) Manga Curta Menino
        4) Mijão
    """
    semi_str = str(semi).lower()

    # Categoria de produto
    if "manga longa" in semi_str:
        categoria = 1
    elif "manga curta" in semi_str and "menina" in semi_str:
        categoria = 2
    elif "manga curta" in semi_str and "menino" in semi_str:
        categoria = 3
    elif "mijão" in semi_str or "mijao" in semi_str:
        categoria = 4
    else:
        categoria = 5

    # Ordem de cores
    if "branco" in semi_str:
        cor_ordem = 1
    elif "off-white" in semi_str or "off white" in semi_str:
        cor_ordem = 2
    elif "rosa" in semi_str:
        cor_ordem = 3
    elif "azul" in semi_str:
        cor_ordem = 4
    elif "vermelho" in semi_str:
        cor_ordem = 5
    elif "marinho" in semi_str:
        cor_ordem = 6
    else:
        cor_ordem = 7

    # Tamanhos
    if "-rn" in semi_str or " rn" in semi_str:
        tamanho_ordem = 1
    elif "-p" in semi_str or " p" in semi_str:
        tamanho_ordem = 2
    elif "-m" in semi_str or " m" in semi_str:
        tamanho_ordem = 3
    elif "-g" in semi_str or " g" in semi_str:
        tamanho_ordem = 4
    else:
        tamanho_ordem = 5

    return categoria, cor_ordem, tamanho_ordem


def explodir_kits(df_vendas_com_mae, df_mae_completa):
    """
    Função principal para "explodir" kits em seus componentes individuais
    (Semi / Gola / Bordado), reaproveitando a estrutura original.

    - df_vendas_com_mae: já mesclado com a planilha mãe.
    - df_mae_completa: planilha mãe completa (códigos → semi/gola/bordado/componentes_codigos).

    Retorna DataFrame com colunas:
      semi, gola, bordado, quantidade
    """
    componentes_finais = []

    df_mae_completa = df_mae_completa.set_index("codigo")

    def obter_componentes(codigo, quantidade):
        lista_componentes_recursiva = []

        try:
            produto = df_mae_completa.loc[codigo]
        except KeyError:
            return []

        # 1. Componente direto (semi/gola/bordado)
        semi_valido = False
        if "semi" in produto.index and pd.notna(produto["semi"]):
            if isinstance(produto["semi"], str) and produto["semi"].strip() != "":
                semi_valido = True

        if semi_valido:
            lista_componentes_recursiva.append(
                {
                    "semi": produto["semi"],
                    "gola": produto["gola"] if pd.notna(produto.get("gola", "")) else "",
                    "bordado": produto["bordado"] if pd.notna(produto.get("bordado", "")) else "",
                    "quantidade": quantidade,
                }
            )

        # 2. Componentes aninhados (kits)
        componentes_codigos_valido = False
        if "componentes_codigos" in produto.index and pd.notna(produto["componentes_codigos"]):
            comp_str = str(produto["componentes_codigos"]).strip()
            if comp_str != "" and comp_str.lower() != "nan":
                componentes_codigos_valido = True

        if componentes_codigos_valido:
            codigos_aninhados = str(produto["componentes_codigos"]).split(";")
            for cod_aninhado in codigos_aninhados:
                cod_aninhado = cod_aninhado.strip()
                if cod_aninhado:
                    lista_componentes_recursiva.extend(obter_componentes(cod_aninhado, quantidade))

        return lista_componentes_recursiva

    for _, venda in df_vendas_com_mae.iterrows():
        componentes_finais.extend(obter_componentes(venda["codigo"], venda["quantidade"]))

    return pd.DataFrame(componentes_finais)


def gerar_excel_formatado(df, nome_aba, agrupar_por_semi=False):
    """
    Gera um arquivo Excel formatado a partir de um DataFrame.
    Usado para todos os relatórios baixados.
    """
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = nome_aba

    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    manga_longa_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    manga_curta_menina_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    manga_curta_menino_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    mijao_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    semi_font = Font(bold=True)
    border = Border(
        left=Side(style="thin"),
        right=Side(style="thin"),
        top=Side(style="thin"),
        bottom=Side(style="thin"),
    )

    if agrupar_por_semi:
        # Layout hierárquico Semi → Golas/Bordados
        headers = ["Item", "Quantidade", "Check"]
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border

        df["gola"] = df["gola"].fillna("")
        df["bordado"] = df["bordado"].fillna("")

        # Agrupa semi + gola/bordado
        relatorio_componentes = df.groupby(["semi", "gola", "bordado"])["quantidade"].sum().reset_index()

        relatorio_componentes[["categoria", "cor_ordem", "tamanho_ordem"]] = relatorio_componentes[
            "semi"
        ].apply(lambda x: pd.Series(get_categoria_ordem(x)))

        relatorio_componentes = relatorio_componentes.sort_values(
            ["categoria", "cor_ordem", "tamanho_ordem", "semi", "gola", "bordado"]
        )

        relatorio_hierarquico = []
        for semi_produto, grupo in relatorio_componentes.groupby("semi"):
            total_semi = grupo["quantidade"].sum()
            categoria = grupo["categoria"].iloc[0]

            relatorio_hierarquico.append(
                {
                    "Item": semi_produto,
                    "Quantidade": total_semi,
                    "Check": "",
                    "categoria": categoria,
                    "is_semi": True,
                }
            )

            for _, row in grupo.iterrows():
                componentes_txt = f"{row['gola']} {row['bordado']}".strip()
                if componentes_txt:
                    relatorio_hierarquico.append(
                        {
                            "Item": f"  {componentes_txt}",
                            "Quantidade": row["quantidade"],
                            "Check": "",
                            "categoria": categoria,
                            "is_semi": False,
                        }
                    )

        row_num = 2
        for item in relatorio_hierarquico:
            is_semi = item["is_semi"]
            categoria = item["categoria"]

            fill_color = None
            if is_semi:
                if categoria == 1:
                    fill_color = manga_longa_fill
                elif categoria == 2:
                    fill_color = manga_curta_menina_fill
                elif categoria == 3:
                    fill_color = manga_curta_menino_fill
                elif categoria == 4:
                    fill_color = mijao_fill

            for col_num, key in enumerate(["Item", "Quantidade", "Check"], 1):
                cell = ws.cell(row=row_num, column=col_num, value=item[key])
                cell.border = border
                if is_semi:
                    if col_num == 1:
                        cell.font = semi_font
                    if fill_color:
                        cell.fill = fill_color
            row_num += 1

        ws.column_dimensions["A"].width = 60
        ws.column_dimensions["B"].width = 12
        ws.column_dimensions["C"].width = 8

    else:
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

        # Ajuste de largura
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except Exception:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width

    wb.save(output)
    output.seek(0)
    return output


# ==============================================================================
# 1) CARREGAMENTO DA PLANILHA MÃE
# ==============================================================================

st.header("📁 1. Configuração Inicial — Planilha Mãe")

st.markdown(
    """
<div class="explicacao-box">
<b>O que é a Planilha Mãe?</b><br>
Ela define a “receita” de cada produto:<br>
• <code>codigo</code> → código de venda<br>
• <code>semi</code> → qual semi esse produto usa<br>
• <code>gola</code> → qual gola esse produto usa<br>
• <code>bordado</code> → qual bordado a gola usa<br>
• <code>componentes_codigos</code> (opcional) → códigos extras que compõem kits<br><br>
Você só precisa carregar essa planilha uma vez por sessão.
</div>
""",
    unsafe_allow_html=True,
)


def carregar_planilha_mae(arquivo):
    """Carrega e valida a planilha mãe, atualizando o session_state."""
    try:
        with st.spinner("Carregando e validando Planilha Mãe..."):
            df = load_excel(arquivo)
            df.columns = df.columns.str.strip().str.replace(" ", "_").str.lower()

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
        st.error(f"Erro ao carregar planilha mãe: {str(e)}")


if st.session_state["planilha_mae_carregada"]:
    st.success(
        f"✅ Planilha Mãe carregada: {len(st.session_state['df_mae'])} produtos cadastrados."
    )
    with st.expander("🔄 Recarregar / Atualizar Planilha Mãe"):
        uploaded_mae_nova = st.file_uploader(
            "Substituir Planilha Mãe atual", type=["xlsx"], key="planilha_mae_nova"
        )
        if uploaded_mae_nova:
            carregar_planilha_mae(uploaded_mae_nova)
else:
    st.info(
        "📋 Para começar, carregue a Planilha Mãe (`codigo`, `semi`, `gola`, `bordado`, `componentes_codigos`)."
    )
    uploaded_mae = st.file_uploader(
        "Carregar Planilha Mãe", type=["xlsx"], key="planilha_mae"
    )
    if uploaded_mae:
        carregar_planilha_mae(uploaded_mae)


# ==============================================================================
# 2) PROCESSAMENTO DE VENDAS + ESTOQUE TEMPLATE
# ==============================================================================

if st.session_state["planilha_mae_carregada"]:
    st.header("📊 2. Processar Vendas do Dia")

    st.markdown(
        """
<div class="explicacao-box">
<b>Como deve ser a planilha de vendas?</b><br>
• Formato: Excel (<code>.xlsx</code>)<br>
• Colunas obrigatórias: <code>código</code> e <code>quantidade</code><br>
• Uma linha por venda / produto.<br><br>
O app vai:<br>
1) Somar as quantidades vendidas por código;<br>
2) Consultar o estoque de produtos prontos na <b>template_estoque</b> (modo leitura);<br>
3) Usar o que já tem pronto em estoque;<br>
4) Só explodir em insumos o que <b>realmente falta produzir</b>.
</div>
""",
        unsafe_allow_html=True,
    )

    uploaded_vendas = st.file_uploader(
        "📈 Planilha de Vendas (diária)", type=["xlsx"], key="vendas"
    )

    if uploaded_vendas:
        df_mae = st.session_state["df_mae"]

        try:
            with st.spinner("Carregando vendas..."):
                df_vendas = load_excel(uploaded_vendas)
                df_vendas.columns = (
                    df_vendas.columns.str.strip().str.replace(" ", "_").str.lower()
                )

                if "código" not in df_vendas.columns or "quantidade" not in df_vendas.columns:
                    st.error("❌ A planilha de vendas deve ter as colunas 'código' e 'quantidade'.")
                    st.stop()

                df_vendas = df_vendas.rename(columns={"código": "codigo"})
                df_vendas["quantidade"] = pd.to_numeric(
                    df_vendas["quantidade"], errors="coerce"
                ).fillna(0).astype(int)

                # Agrupa por código (total vendido no período)
                df_vendas_agr = (
                    df_vendas.groupby("codigo", as_index=False)["quantidade"].sum()
                )

            # ------------------------------------------------------------------
            # 2.1 Ler template_estoque e cruzar com produtos prontos
            # ------------------------------------------------------------------
            st.subheader("📦 Situação dos Produtos Prontos (template_estoque)")

            df_estoque = carregar_template_estoque()
            if df_estoque.empty:
                st.warning(
                    "⚠️ Não foi possível carregar a template_estoque. "
                    "O app vai considerar que não há produto pronto em estoque."
                )
                df_estoque_produtos = pd.DataFrame(
                    columns=["codigo", "nome", "estoque_atual"]
                )
            else:
                # Considera tudo como "produto pronto" para esse nível
                df_estoque_produtos = df_estoque[["codigo", "nome", "estoque_atual"]].copy()

            df_merge_prod = df_vendas_agr.merge(
                df_estoque_produtos, on="codigo", how="left"
            )
            df_merge_prod["estoque_atual"] = df_merge_prod["estoque_atual"].fillna(0)

            df_merge_prod["usando_estoque_pronto"] = df_merge_prod[
                ["quantidade", "estoque_atual"]
            ].min(axis=1)
            df_merge_prod["faltante_produto"] = (
                df_merge_prod["quantidade"] - df_merge_prod["estoque_atual"]
            )
            df_merge_prod["faltante_produto"] = df_merge_prod["faltante_produto"].clip(
                lower=0
            )

            st.markdown(
                """
<div class="explicacao-box">
<b>O que você está vendo aqui?</b><br>
• <b>quantidade</b> → total vendido no período;<br>
• <b>estoque_atual</b> → quanto já existe pronto na template_estoque;<br>
• <b>faltante_produto</b> → quanto ainda precisa ser produzido;<br><br>
<b>Somente os códigos com faltante_produto &gt; 0 serão explodidos em insumos.</b>
</div>
""",
                unsafe_allow_html=True,
            )

            tabela_prod = df_merge_prod[["codigo", "nome", "quantidade", "estoque_atual", "faltante_produto"]]
            st.dataframe(tabela_prod, use_container_width=True, height=350)

            # Download da situação de produtos prontos
            excel_produtos_prontos = gerar_excel_formatado(
                tabela_prod, "produtos_prontos", agrupar_por_semi=False
            )
            st.download_button(
                "📥 Baixar situação de produtos prontos (Excel)",
                excel_produtos_prontos,
                "situacao_produtos_prontos.xlsx",
            )
            st.caption("Esse arquivo mostra tudo o que foi vendido x o que já tem pronto x o que falta produzir.")

            # ------------------------------------------------------------------
            # 2.2 Filtra apenas faltantes para explodir em insumos
            # ------------------------------------------------------------------
            df_faltantes = df_merge_prod[df_merge_prod["faltante_produto"] > 0].copy()

            if df_faltantes.empty:
                st.markdown(
                    """
<div class="sucesso-box">
✅ Todas as vendas foram cobertas com estoque de produtos prontos da template_estoque.<br>
Não há necessidade de explodir insumos hoje.
</div>
""",
                    unsafe_allow_html=True,
                )
                st.stop()

            # Usa apenas a quantidade faltante
            df_faltantes = df_faltantes[["codigo", "faltante_produto"]].rename(
                columns={"faltante_produto": "quantidade"}
            )

            # Mescla com planilha mãe para ter semi/gola/bordado
            df_mae_cols = df_mae.copy()
            df_mae_cols.columns = df_mae_cols.columns.str.lower()
            df_merged = df_faltantes.merge(df_mae_cols, on="codigo", how="left")

            codigos_sem_mae = df_merged[df_merged["semi"].isna()]["codigo"].unique()
            dados_validos_df = df_merged.dropna(subset=["semi"])

            if len(codigos_sem_mae) > 0:
                st.markdown(
                    """
<div class="alerta-box">
⚠️ Existem códigos nas vendas que <b>não estão na Planilha Mãe</b>.<br>
Esses códigos não serão explodidos em insumos até que sejam cadastrados.
</div>
""",
                    unsafe_allow_html=True,
                )
                df_faltantes_mae = pd.DataFrame({"codigo": codigos_sem_mae})
                excel_faltantes_mae = gerar_excel_formatado(
                    df_faltantes_mae, "codigos_sem_mae", agrupar_por_semi=False
                )
                st.download_button(
                    "📥 Baixar lista de códigos sem Planilha Mãe",
                    excel_faltantes_mae,
                    "codigos_sem_planilha_mae.xlsx",
                )
                st.caption("Use este arquivo para completar a Planilha Mãe com semi / gola / bordado.")

            if dados_validos_df.empty:
                st.error("Não há nenhum código faltante com semi configurado na Planilha Mãe.")
                st.stop()

            # ------------------------------------------------------------------
            # 2.3 Explode insumos (apenas faltantes) → Semi / Gola / Bordado
            # ------------------------------------------------------------------
            st.subheader("🧵 Explosão em Insumos (apenas do que falta produzir)")

            with st.spinner("Explodindo kits e gerando insumos..."):
                dados_explodidos = explodir_kits(dados_validos_df, df_mae_cols)

            if dados_explodidos.empty:
                st.warning("Nenhum insumo foi encontrado para os códigos faltantes.")
                st.stop()

            st.markdown(
                """
<div class="explicacao-box">
<b>O que é essa tabela?</b><br>
Cada linha representa um insumo (Semi, Gola, Bordado) necessário para cobrir apenas o que <b>não</b> foi atendido com produto pronto.<br>
A coluna <code>quantidade</code> já considera o total de peças faltantes.
</div>
""",
                unsafe_allow_html=True,
            )

            st.dataframe(dados_explodidos, use_container_width=True, height=300)

            # ------------------------------------------------------------------
            # 2.4 Cruzar insumos com estoque da template_estoque
            # ------------------------------------------------------------------
            st.subheader("🏭 Planejamento de Produção por Semi / Gola / Bordado")

            st.markdown(
                """
<div class="explicacao-box">
<b>Agora o app cruza os insumos necessários com o estoque da template_estoque:</b><br>
1) Verifica se existe <b>Semi</b> em estoque;<br>
2) Verifica se existem <b>Golas</b> em estoque;<br>
3) Se faltar gola, calcula <b>Bordados</b> necessários para completar as golas faltantes.<br><br>
Resultado:
• Você vê exatamente <b>o que precisa produzir hoje</b>, organizado por Semi → Golas → Bordados.
</div>
""",
                unsafe_allow_html=True,
            )

            # Dicionários de estoque por nome (Semi / Gola / Bordado)
            if not df_estoque.empty:
                # Normaliza nomes
                df_estoque["nome_norm"] = df_estoque["nome"].astype(str).str.strip().str.lower()
                df_estoque["categoria_norm"] = df_estoque["categoria"].astype(str).str.strip().str.lower()

                def build_dict(cat):
                    sub = df_estoque[df_estoque["categoria_norm"] == cat].copy()
                    return dict(zip(sub["nome_norm"], sub["estoque_atual"]))

                estoque_semi_dict = build_dict("semi")
                estoque_gola_dict = build_dict("gola")
                estoque_bordado_dict = build_dict("bordado")
            else:
                estoque_semi_dict = {}
                estoque_gola_dict = {}
                estoque_bordado_dict = {}

            # ---- Semis agregados ----
            semi_agg = (
                dados_explodidos.groupby("semi")["quantidade"].sum().reset_index()
            )
            semi_agg["semi_norm"] = semi_agg["semi"].astype(str).str.strip().str.lower()
            semi_agg["estoque_atual"] = semi_agg["semi_norm"].map(estoque_semi_dict).fillna(0)
            semi_agg["faltante_semi"] = (
                semi_agg["quantidade"] - semi_agg["estoque_atual"]
            ).clip(lower=0)

            semi_agg[["categoria", "cor_ordem", "tamanho_ordem"]] = semi_agg["semi"].apply(
                lambda x: pd.Series(get_categoria_ordem(x))
            )
            semi_agg_sorted = semi_agg.sort_values(
                ["categoria", "cor_ordem", "tamanho_ordem", "semi"]
            ).reset_index(drop=True)

            # ---- Golas agregadas por Semi+Gola ----
            dados_explodidos["gola"] = dados_explodidos["gola"].fillna("")
            golas = dados_explodidos[dados_explodidos["gola"].str.strip() != ""].copy()

            if not golas.empty:
                gola_agg = (
                    golas.groupby(["semi", "gola"])["quantidade"].sum().reset_index()
                )
                gola_agg["gola_norm"] = gola_agg["gola"].astype(str).str.strip().str.lower()
                gola_agg["estoque_atual"] = gola_agg["gola_norm"].map(estoque_gola_dict).fillna(0)
                gola_agg["faltante_gola"] = (
                    gola_agg["quantidade"] - gola_agg["estoque_atual"]
                ).clip(lower=0)
            else:
                gola_agg = pd.DataFrame(columns=["semi", "gola", "quantidade", "estoque_atual", "faltante_gola"])

            # ---- Bordados (apenas quando faltar gola) ----
            dados_explodidos["bordado"] = dados_explodidos["bordado"].fillna("")
            bordados_list = []

            if not gola_agg.empty:
                # Para cada combinação semi+gola com falta, usa o bordado correspondente
                falta_gola_df = gola_agg[gola_agg["faltante_gola"] > 0]
                if not falta_gola_df.empty:
                    # Mapeia (semi,gola) -> bordado mais comum
                    mapa_bordado = (
                        dados_explodidos.groupby(["semi", "gola"])["bordado"]
                        .agg(lambda x: x.mode().iloc[0] if len(x.mode()) > 0 else "")
                        .reset_index()
                    )
                    falta_gola_df = falta_gola_df.merge(
                        mapa_bordado, on=["semi", "gola"], how="left"
                    )

                    for _, row in falta_gola_df.iterrows():
                        qtd_bordados = row["faltante_gola"]
                        if qtd_bordados <= 0:
                            continue
                        bordado_nome = str(row["bordado"]).strip()
                        if bordado_nome == "":
                            continue

                        bordados_list.append(
                            {
                                "bordado": bordado_nome,
                                "quantidade_necessaria": qtd_bordados,
                            }
                        )

            if bordados_list:
                bordados_df = pd.DataFrame(bordados_list)
                bordados_agg = (
                    bordados_df.groupby("bordado")["quantidade_necessaria"]
                    .sum()
                    .reset_index()
                )
                bordados_agg = bordados_agg.rename(
                    columns={"quantidade_necessaria": "quantidade"}
                )
                bordados_agg["bordado_norm"] = (
                    bordados_agg["bordado"].astype(str).str.strip().str.lower()
                )
                bordados_agg["estoque_atual"] = bordados_agg["bordado_norm"].map(
                    estoque_bordado_dict
                ).fillna(0)
                bordados_agg["faltante_bordado"] = (
                    bordados_agg["quantidade"] - bordados_agg["estoque_atual"]
                ).clip(lower=0)
            else:
                bordados_agg = pd.DataFrame(
                    columns=["bordado", "quantidade", "estoque_atual", "faltante_bordado"]
                )

            # ------------------------------------------------------------------
            # 2.5 Relatório hierárquico na tela
            # ------------------------------------------------------------------
            st.markdown("### 📌 Visão por Semi → Golas → Bordados")

            st.markdown(
                """
<div class="explicacao-box">
<b>Como ler essa seção?</b><br>
• Cada bloco começa com um <b>Semi</b> e a quantidade total necessária;<br>
• Abaixo, aparecem as <b>Golas</b> associadas àquele semi e o quanto falta;<br>
• Se faltar gola, o sistema calcula automaticamente os <b>Bordados</b> necessários.
</div>
""",
                unsafe_allow_html=True,
            )

            for _, semi_row in semi_agg_sorted.iterrows():
                semi_nome = semi_row["semi"]
                qtd_semi = int(semi_row["quantidade"])
                est_semi = int(semi_row["estoque_atual"])
                falt_semi = int(semi_row["faltante_semi"])

                st.markdown(f"#### 🧵 Semi: **{semi_nome}**")
                st.write(
                    f"• Necessário: **{qtd_semi}** | Em estoque (Semi): **{est_semi}** | Faltando Semi: **{falt_semi}**"
                )

                # Golas deste semi
                sub_gola = gola_agg[gola_agg["semi"] == semi_nome]
                if not sub_gola.empty:
                    st.write("**Golas para este Semi:**")
                    gola_show = sub_gola[["gola", "quantidade", "estoque_atual", "faltante_gola"]].copy()
                    gola_show.columns = [
                        "Gola",
                        "Qtd Necessária",
                        "Estoque Atual (Gola)",
                        "Faltante Gola",
                    ]
                    st.dataframe(gola_show, use_container_width=True, height=180)
                else:
                    st.write("_Nenhuma gola específica cadastrada para este Semi na Planilha Mãe._")

                st.markdown("---")

            # ------------------------------------------------------------------
            # 2.6 Download dos relatórios finais
            # ------------------------------------------------------------------
            st.subheader("📥 3. Relatórios para Download")

            col_r1, col_r2, col_r3, col_r4 = st.columns(4)

            # a) Relatório Componentes (hierárquico Semi > gola/bordado)
            with col_r1:
                excel_componentes = gerar_excel_formatado(
                    dados_explodidos, "Componentes_por_Semi", agrupar_por_semi=True
                )
                st.download_button(
                    "📋 Componentes por Semi (Excel)",
                    excel_componentes,
                    "componentes_por_semi.xlsx",
                    key="btn_comp_semi",
                )
                st.caption(
                    "Semi na linha principal e, logo abaixo, as golas/bordados com as quantidades necessárias."
                )

            # b) Resumo de Semis
            with col_r2:
                semi_res = semi_agg_sorted[["semi", "quantidade", "estoque_atual", "faltante_semi"]].copy()
                semi_res.columns = [
                    "Semi",
                    "Qtd Necessária",
                    "Estoque Atual (Semi)",
                    "Faltante Semi",
                ]
                excel_semis = gerar_excel_formatado(
                    semi_res, "Resumo_Semis", agrupar_por_semi=False
                )
                st.download_button(
                    "🧵 Resumo de Semis (Excel)",
                    excel_semis,
                    "resumo_semis_producao.xlsx",
                    key="btn_semis",
                )
                st.caption("Lista todos os semis, o quanto precisa, o que já tem e o que falta produzir.")

            # c) Resumo de Golas
            with col_r3:
                if not gola_agg.empty:
                    gola_res = gola_agg[["gola", "quantidade", "estoque_atual", "faltante_gola"]].copy()
                    gola_res.columns = [
                        "Gola",
                        "Qtd Necessária",
                        "Estoque Atual (Gola)",
                        "Faltante Gola",
                    ]
                    excel_golas = gerar_excel_formatado(
                        gola_res, "Resumo_Golas", agrupar_por_semi=False
                    )
                    st.download_button(
                        "👔 Resumo de Golas (Excel)",
                        excel_golas,
                        "resumo_golas_producao.xlsx",
                        key="btn_golas",
                    )
                    st.caption("Quais golas você precisa hoje, quanto tem e quanto falta fazer.")
                else:
                    st.info("Não há golas mapeadas para esse conjunto de vendas.")

            # d) Resumo de Bordados
            with col_r4:
                if not bordados_agg.empty:
                    bord_res = bordados_agg[
                        ["bordado", "quantidade", "estoque_atual", "faltante_bordado"]
                    ].copy()
                    bord_res.columns = [
                        "Bordado",
                        "Qtd Necessária",
                        "Estoque Atual (Bordado)",
                        "Faltante Bordado",
                    ]
                    excel_bordados = gerar_excel_formatado(
                        bord_res, "Resumo_Bordados", agrupar_por_semi=False
                    )
                    st.download_button(
                        "🎨 Resumo de Bordados (Excel)",
                        excel_bordados,
                        "resumo_bordados_producao.xlsx",
                        key="btn_bordados",
                    )
                    st.caption(
                        "Somente bordados necessários para cobrir as golas que estão faltando."
                    )
                else:
                    st.info("Nenhum bordado adicional foi necessário para esta produção.")

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante o processamento: {str(e)}")


# ==============================================================================
# SIDEBAR - AJUDA RÁPIDA
# ==============================================================================

st.sidebar.markdown("---")
st.sidebar.info(
    "💡 A Planilha Mãe permanece carregada apenas nesta sessão. "
    "Se fechar o navegador, será preciso carregá-la novamente."
)
st.sidebar.markdown("---")
st.sidebar.info(
    "📦 A template_estoque é acessada em modo SOMENTE LEITURA.\n\n"
    "• O app de relatórios <b>nunca</b> altera a planilha de estoque.\n"
    "• Quem altera estoque é apenas o cockpit <code>estoque-completo-v3</code>."
)
