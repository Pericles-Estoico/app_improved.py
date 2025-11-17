# app_improved.py
import streamlit as st
import pandas as pd
from io import BytesIO, StringIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import numpy as np
import requests

# ==============================================================================
# CONFIGURAÇÕES E ESTILOS
# ==============================================================================

st.set_page_config(page_title="Pure & Posh Baby - Sistema de Relatórios", page_icon="👑", layout="wide")

# 🔗 URL DA PLANILHA DE ESTOQUE (template_estoque) – APENAS LEITURA
# Agora usando export?format=csv&gid=1456159896 (mesmo gid do link que você mandou)
TEMPLATE_ESTOQUE_URL = (
    "https://docs.google.com/spreadsheets/d/"
    "1PpiMQingHf4llA03BiPIuPJPIZqul4grRU_emWDEK1o"
    "/export?format=csv&gid=1456159896"
)

# Header visual
st.markdown("""
<style>
.centered-title { text-align: center; width: 100%; margin: 0 auto; }
@media (max-width: 768px) { .centered-title { text-align: center; } }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios de Vendas v4.1")
st.markdown("**Pure & Posh Baby**")
st.markdown('</div>', unsafe_allow_html=True)

# Session state inicial
if 'planilha_mae_carregada' not in st.session_state:
    st.session_state['planilha_mae_carregada'] = False
if 'df_mae' not in st.session_state:
    st.session_state['df_mae'] = None

# ==============================================================================
# FUNÇÕES CORE
# ==============================================================================

@st.cache_data
def load_excel(arquivo):
    """Carrega um arquivo Excel em um DataFrame, com cache para performance."""
    return pd.read_excel(arquivo)

@st.cache_data(ttl=60)
def carregar_template_estoque_raw():
    """
    Lê a aba 'template_estoque' via gid, apenas leitura.
    """
    try:
        r = requests.get(TEMPLATE_ESTOQUE_URL, timeout=15)
        r.raise_for_status()
        df = pd.read_csv(StringIO(r.text))
        if df.empty:
            return None
        # normaliza colunas
        df.columns = df.columns.str.strip().str.lower()
        return df
    except Exception as e:
        st.warning(f"⚠️ Erro ao ler template_estoque: {e}")
        return None

def get_categoria_ordem(semi):
    """
    Determina a categoria e a ordem de um item 'semi' para ordenação nos relatórios.
    
    ORDEM:
      1 = Manga Longa
      2 = Manga Curta Menina
      3 = Manga Curta Menino
      4 = Mijão
      5 = Outros
    
    Dentro de cada categoria:
      Cor: Branco, Off, Rosa, Azul, Vermelho, Marinho, Outras
      Tamanho: RN, P, M, G, 1, 2, 3, 4, Outros
    """
    semi_str = str(semi).lower()
    
    # Categoria
    if 'manga longa' in semi_str:
        categoria = 1
    elif 'manga curta' in semi_str and 'menina' in semi_str:
        categoria = 2
    elif 'manga curta' in semi_str and 'menino' in semi_str:
        categoria = 3
    elif 'mijão' in semi_str or 'mijao' in semi_str:
        categoria = 4
    else:
        categoria = 5
    
    # Cor
    if 'branco' in semi_str:
        cor_ordem = 1
    elif 'off-white' in semi_str or 'off white' in semi_str:
        cor_ordem = 2
    elif 'rosa' in semi_str:
        cor_ordem = 3
    elif 'azul' in semi_str:
        cor_ordem = 4
    elif 'vermelho' in semi_str:
        cor_ordem = 5
    elif 'marinho' in semi_str:
        cor_ordem = 6
    else:
        cor_ordem = 7
    
    # Tamanho
    if '-rn' in semi_str or ' rn' in semi_str:
        tamanho_ordem = 1
    elif '-p' in semi_str or ' p' in semi_str:
        tamanho_ordem = 2
    elif '-m' in semi_str or ' m' in semi_str:
        tamanho_ordem = 3
    elif '-g' in semi_str or ' g' in semi_str:
        tamanho_ordem = 4
    elif '-1' in semi_str or ' 1' in semi_str:
        tamanho_ordem = 5
    elif '-2' in semi_str or ' 2' in semi_str:
        tamanho_ordem = 6
    elif '-3' in semi_str or ' 3' in semi_str:
        tamanho_ordem = 7
    elif '-4' in semi_str or ' 4' in semi_str:
        tamanho_ordem = 8
    else:
        tamanho_ordem = 9
    
    return categoria, cor_ordem, tamanho_ordem

def explodir_kits(df_vendas_com_mae, df_mae_completa):
    """
    Explode kits + produtos em componentes individuais (Semi / Gola / Bordado),
    multiplicando pela quantidade NECESSÁRIA (faltante de produto pronto).
    """
    componentes_finais = []
    
    # Índice por código para acesso rápido
    df_mae_completa = df_mae_completa.set_index('codigo')

    def obter_componentes(codigo, quantidade):
        """Recursivo: encontra todos os componentes (semi/gola/bordado + filhos) de um código."""
        lista_componentes_recursiva = []
        
        try:
            produto = df_mae_completa.loc[codigo]
        except KeyError:
            # Código não existe na planilha mãe
            return []

        # 1. Componentes diretos (semi/gola/bordado)
        semi_valido = False
        if 'semi' in produto.index and pd.notna(produto['semi']):
            if isinstance(produto['semi'], str) and produto['semi'].strip() != '':
                semi_valido = True

        if semi_valido:
            lista_componentes_recursiva.append({
                'semi': produto['semi'],
                'gola': produto['gola'] if ('gola' in produto.index and pd.notna(produto['gola'])) else '',
                'bordado': produto['bordado'] if ('bordado' in produto.index and pd.notna(produto['bordado'])) else '',
                'quantidade': quantidade
            })

        # 2. Kits (componentes_codigos)
        if 'componentes_codigos' in produto.index and pd.notna(produto['componentes_codigos']):
            componentes_str = str(produto['componentes_codigos']).strip()
            if componentes_str != '' and componentes_str.lower() != 'nan':
                codigos_aninhados = componentes_str.split(';')
                for cod_aninhado in codigos_aninhados:
                    cod_aninhado = cod_aninhado.strip()
                    if cod_aninhado:
                        lista_componentes_recursiva.extend(obter_componentes(cod_aninhado, quantidade))
        
        return lista_componentes_recursiva

    for _, venda in df_vendas_com_mae.iterrows():
        # venda['quantidade'] aqui já é a quantidade FALTANTE de produto pronto
        componentes_finais.extend(obter_componentes(venda['codigo'], venda['quantidade']))

    return pd.DataFrame(componentes_finais)

def gerar_excel_formatado(df, nome_arquivo, agrupar_por_semi=False):
    """
    Gera um arquivo Excel formatado a partir de um DataFrame.

    Se agrupar_por_semi=True:
      - Agrupa por (semi, gola, bordado)
      - Ordena por categoria/cor/tamanho
      - Cria blocos:
          SEMI (linha pai)
            - golas necessárias abaixo
            - se não houver gola -> bordado abaixo
    """
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"
    
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    manga_longa_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
    manga_curta_menina_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    manga_curta_menino_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    mijao_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    semi_font = Font(bold=True)
    border = Border(left=Side(style='thin'), right=Side(style='thin'),
                    top=Side(style='thin'), bottom=Side(style='thin'))
    
    if agrupar_por_semi:
        # Cabeçalho
        headers = ['Item', 'Quantidade', 'Check']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
        
        # Limpar nulos
        df = df.copy()
        df['gola'] = df['gola'].fillna('')
        df['bordado'] = df['bordado'].fillna('')
        
        # Agrupa por semi + gola + bordado
        relatorio_componentes = df.groupby(['semi', 'gola', 'bordado'])['quantidade'].sum().reset_index()
        
        # Ordenação
        relatorio_componentes[['categoria', 'cor_ordem', 'tamanho_ordem']] = relatorio_componentes['semi'].apply(
            lambda x: pd.Series(get_categoria_ordem(x))
        )
        relatorio_componentes = relatorio_componentes.sort_values(
            ['categoria', 'cor_ordem', 'tamanho_ordem', 'semi', 'gola', 'bordado']
        )
        
        # Estrutura hierárquica
        relatorio_hierarquico = []
        for semi_produto, grupo in relatorio_componentes.groupby('semi'):
            total_semi = grupo['quantidade'].sum()
            categoria = grupo['categoria'].iloc[0]
            
            # Linha SEMI
            relatorio_hierarquico.append({
                'Item': semi_produto,
                'Quantidade': total_semi,
                'Check': '',
                'categoria': categoria,
                'is_semi': True
            })
            
            # Linhas de componentes (golas/bordados)
            for _, row in grupo.iterrows():
                if row['gola'] and str(row['gola']).strip() != '':
                    componente = f"{row['gola']}"
                elif row['bordado'] and str(row['bordado']).strip() != '':
                    componente = f"{row['bordado']}"
                else:
                    continue
                
                relatorio_hierarquico.append({
                    'Item': f"  {componente}",
                    'Quantidade': row['quantidade'],
                    'Check': '',
                    'categoria': categoria,
                    'is_semi': False
                })

        # Escreve
        row_num = 2
        for item in relatorio_hierarquico:
            is_semi = item['is_semi']
            categoria = item['categoria']
            
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
            
            for col_num, key in enumerate(['Item', 'Quantidade', 'Check'], 1):
                cell = ws.cell(row=row_num, column=col_num, value=item[key])
                cell.border = border
                if is_semi:
                    if col_num == 1:
                        cell.font = semi_font
                    if fill_color:
                        cell.fill = fill_color
            row_num += 1
        
        ws.column_dimensions['A'].width = 60
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 8
        
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
        
        # Autoajuste
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value is not None and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output)
    output.seek(0)
    return output

def cruzar_com_estoque_insumos(dados_explodidos, df_estoque_raw):
    """
    Cruzamento com a planilha template_estoque para insumos:
      - Semis
      - Golas
      - Bordados (apenas quando não há gola)
    """
    if df_estoque_raw is None or df_estoque_raw.empty:
        return None, None, None

    df_est = df_estoque_raw.copy()
    df_est.columns = df_est.columns.str.strip().str.lower()
    if 'nome' not in df_est.columns or 'estoque_atual' not in df_est.columns:
        return None, None, None

    df_est['item'] = df_est['nome'].astype(str).str.strip().str.lower()
    df_est['estoque_atual'] = pd.to_numeric(df_est['estoque_atual'], errors='coerce').fillna(0)

    # --- SEMIS ---
    semis = dados_explodidos.copy()
    semis = semis.groupby('semi', as_index=False)['quantidade'].sum()
    semis = semis[semis['semi'].notna() & (semis['semi'].astype(str).str.strip() != '')]
    semis['semi_key'] = semis['semi'].astype(str).str.strip().str.lower()
    semis[['categoria', 'cor_ordem', 'tamanho_ordem']] = semis['semi'].apply(
        lambda x: pd.Series(get_categoria_ordem(x))
    )

    semis = semis.merge(
        df_est[['item', 'estoque_atual']],
        how='left',
        left_on='semi_key',
        right_on='item'
    )
    semis['estoque_atual'] = semis['estoque_atual'].fillna(0)
    semis['falta'] = (semis['quantidade'] - semis['estoque_atual']).clip(lower=0)
    semis_falta = semis[semis['falta'] > 0].copy()
    semis_falta = semis_falta.sort_values(['categoria', 'cor_ordem', 'tamanho_ordem', 'semi'])

    semis_falta = semis_falta[['semi', 'quantidade', 'estoque_atual', 'falta']]
    semis_falta.columns = ['Semi', 'Qtd Necessária', 'Estoque Atual', 'Falta']

    # --- GOLAS ---
    golas = dados_explodidos.copy()
    golas = golas[golas['gola'].notna() & (golas['gola'].astype(str).str.strip() != '')]
    if not golas.empty:
        golas = golas.groupby('gola', as_index=False)['quantidade'].sum()
        golas['gola_key'] = golas['gola'].astype(str).str.strip().str.lower()

        golas = golas.merge(
            df_est[['item', 'estoque_atual']],
            how='left',
            left_on='gola_key',
            right_on='item'
        )
        golas['estoque_atual'] = golas['estoque_atual'].fillna(0)
        golas['falta'] = (golas['quantidade'] - golas['estoque_atual']).clip(lower=0)
        golas_falta = golas[golas['falta'] > 0].copy()
        golas_falta = golas_falta.sort_values(['gola'])

        golas_falta = golas_falta[['gola', 'quantidade', 'estoque_atual', 'falta']]
        golas_falta.columns = ['Gola', 'Qtd Necessária', 'Estoque Atual', 'Falta']
    else:
        golas_falta = pd.DataFrame(columns=['Gola', 'Qtd Necessária', 'Estoque Atual', 'Falta'])

    # --- BORDADOS (sem gola) ---
    bords = dados_explodidos.copy()
    bords = bords[
        (bords['gola'].isna() | (bords['gola'].astype(str).str.strip() == '')) &
        (bords['bordado'].notna()) &
        (bords['bordado'].astype(str).str.strip() != '')
    ]
    if not bords.empty:
        bords = bords.groupby('bordado', as_index=False)['quantidade'].sum()
        bords['bordado_key'] = bords['bordado'].astype(str).str.strip().str.lower()

        bords = bords.merge(
            df_est[['item', 'estoque_atual']],
            how='left',
            left_on='bordado_key',
            right_on='item'
        )
        bords['estoque_atual'] = bords['estoque_atual'].fillna(0)
        bords['falta'] = (bords['quantidade'] - bords['estoque_atual']).clip(lower=0)
        bords_falta = bords[bords['falta'] > 0].copy()
        bords_falta = bords_falta.sort_values(['bordado'])

        bords_falta = bords_falta[['bordado', 'quantidade', 'estoque_atual', 'falta']]
        bords_falta.columns = ['Bordado', 'Qtd Necessária', 'Estoque Atual', 'Falta']
    else:
        bords_falta = pd.DataFrame(columns=['Bordado', 'Qtd Necessária', 'Estoque Atual', 'Falta'])

    return semis_falta, golas_falta, bords_falta

# ==============================================================================
# INTERFACE DO STREAMLIT
# ==============================================================================

st.header("📁 Configuração Inicial")

def carregar_planilha_mae(arquivo):
    """Carrega e valida a Planilha Mãe."""
    try:
        with st.spinner("Carregando e validando Planilha Mãe..."):
            df = load_excel(arquivo)
            df.columns = df.columns.str.strip().str.replace(" ", "_").str.lower()
            
            colunas_essenciais = ['codigo', 'semi', 'gola', 'bordado']
            if not all(col in df.columns for col in colunas_essenciais):
                st.error(f"❌ Erro: A Planilha Mãe deve conter as colunas: {', '.join(colunas_essenciais)}.")
                return False

            if 'componentes_codigos' not in df.columns:
                df['componentes_codigos'] = ''

            st.session_state['df_mae'] = df
            st.session_state['planilha_mae_carregada'] = True
            st.success(f"✅ Planilha Mãe carregada: {len(df)} produtos cadastrados.")
            st.rerun()
    except Exception as e:
        st.error(f"Erro ao carregar planilha mãe: {str(e)}")

# Upload da Planilha Mãe
if st.session_state['planilha_mae_carregada']:
    st.success(f"✅ Planilha Mãe carregada: {len(st.session_state['df_mae'])} produtos cadastrados.")
    with st.expander("🔄 Recarregar/Atualizar Planilha Mãe"):
        uploaded_mae_nova = st.file_uploader("Substitua a Planilha Mãe atual", type=["xlsx"], key="planilha_mae_nova")
        if uploaded_mae_nova:
            carregar_planilha_mae(uploaded_mae_nova)
else:
    st.info("📋 Para começar, carregue a Planilha Mãe.\n\nEla deve conter as colunas: `codigo`, `semi`, `gola`, `bordado` e, opcionalmente, `componentes_codigos` para kits.")
    uploaded_mae = st.file_uploader("Carregar Planilha Mãe", type=["xlsx"], key="planilha_mae")
    if uploaded_mae:
        carregar_planilha_mae(uploaded_mae)

# Carrega estoque online (produtos prontos + insumos)
df_estoque_raw = carregar_template_estoque_raw()
if df_estoque_raw is None:
    st.warning("⚠️ Não foi possível carregar a aba `template_estoque`. "
               "A lógica de falta de produto pronto será ignorada e o app trabalhará somente com as vendas.")
else:
    st.success(f"📦 Estoque (`template_estoque`) carregado ({len(df_estoque_raw)} linhas).")

# --- Processamento de Vendas ---
if st.session_state['planilha_mae_carregada']:
    st.header("📊 Processamento Diário de Vendas")
    
    st.markdown("""
    **Formato esperado da planilha de Vendas (diária)**  
    - Arquivo: `.xlsx`  
    - Colunas mínimas:
        - `código` (ou `codigo`)
        - `quantidade`
    """)

    uploaded_vendas = st.file_uploader("📈 Planilha de Vendas (diária)", type=["xlsx"], key="vendas")
    
    if uploaded_vendas:
        try:
            with st.spinner("Processando vendas..."):
                df_vendas = load_excel(uploaded_vendas)
                df_vendas.columns = df_vendas.columns.str.strip().str.replace(' ', '_').str.lower()

                if 'código' not in df_vendas.columns and 'codigo' not in df_vendas.columns:
                    st.error("❌ Planilha de vendas deve ter coluna 'código' ou 'codigo'.")
                    st.stop()
                if 'quantidade' not in df_vendas.columns:
                    st.error("❌ Planilha de vendas deve ter coluna 'quantidade'.")
                    st.stop()

                if 'código' in df_vendas.columns:
                    df_vendas = df_vendas.rename(columns={'código': 'codigo'})
                
                df_vendas['quantidade'] = pd.to_numeric(df_vendas['quantidade'], errors='coerce').fillna(0).astype(int)
                df_vendas = df_vendas[df_vendas['quantidade'] > 0]

                df_mae = st.session_state['df_mae']

                # ------------------------------------------------------------------
                # 1) VERIFICA ESTOQUE DE PRODUTO PRONTO (template_estoque)
                # ------------------------------------------------------------------
                if df_estoque_raw is not None:
                    df_est_prod = df_estoque_raw.copy()
                    # tenta achar coluna de código
                    col_codigo_est = None
                    for cand in ['codigo', 'código']:
                        if cand in df_est_prod.columns:
                            col_codigo_est = cand
                            break

                    if col_codigo_est is not None and 'estoque_atual' in df_est_prod.columns:
                        df_est_prod = df_est_prod[[col_codigo_est, 'nome', 'estoque_atual']].copy()
                        df_est_prod.rename(columns={col_codigo_est: 'codigo'}, inplace=True)
                        df_est_prod['estoque_atual'] = pd.to_numeric(df_est_prod['estoque_atual'], errors='coerce').fillna(0)

                        df_merge_prod = pd.merge(
                            df_vendas[['codigo', 'quantidade']],
                            df_est_prod,
                            on='codigo',
                            how='left'
                        )

                        df_merge_prod['estoque_atual'] = df_merge_prod['estoque_atual'].fillna(0)
                        df_merge_prod['faltante_produto'] = (df_merge_prod['quantidade'] - df_merge_prod['estoque_atual']).clip(lower=0)

                        st.subheader("📦 Situação dos Produtos Prontos (template_estoque)")
                        st.dataframe(
                            df_merge_prod[['codigo', 'nome', 'quantidade', 'estoque_atual', 'faltante_produto']],
                            use_container_width=True
                        )

                        df_para_produzir = df_merge_prod[df_merge_prod['faltante_produto'] > 0].copy()
                        df_para_produzir = df_para_produzir[['codigo', 'faltante_produto']].rename(columns={'faltante_produto': 'quantidade'})

                        if df_para_produzir.empty:
                            st.success("✅ Todos os produtos vendidos têm estoque suficiente na `template_estoque`. Nada a produzir hoje.")
                            df_merge_mae = pd.merge(df_vendas[['codigo', 'quantidade']], df_mae, on='codigo', how='left')
                            codigos_faltantes = df_merge_mae[df_merge_mae['semi'].isna()]['codigo'].unique()
                            dados_validos_df = df_merge_mae.dropna(subset=['semi'])
                        else:
                            st.info(f"⚙️ {len(df_para_produzir)} código(s) com falta de produto pronto. Apenas esses serão explodidos em insumos.")
                            df_merge_mae = pd.merge(df_para_produzir, df_mae, on='codigo', how='left')
                            codigos_faltantes = df_merge_mae[df_merge_mae['semi'].isna()]['codigo'].unique()
                            dados_validos_df = df_merge_mae.dropna(subset=['semi'])
                    else:
                        st.warning("⚠️ Não encontrei colunas 'codigo'/'código' + 'estoque_atual' em template_estoque. "
                                   "Voltando para modo simples (explode todas as vendas).")
                        df_merge_mae = pd.merge(df_vendas[['codigo', 'quantidade']], df_mae, on='codigo', how='left')
                        codigos_faltantes = df_merge_mae[df_merge_mae['semi'].isna()]['codigo'].unique()
                        dados_validos_df = df_merge_mae.dropna(subset=['semi'])
                else:
                    # Sem estoque online → comportamento antigo (explode tudo que vendeu)
                    df_merge_mae = pd.merge(df_vendas[['codigo', 'quantidade']], df_mae, on='codigo', how='left')
                    codigos_faltantes = df_merge_mae[df_merge_mae['semi'].isna()]['codigo'].unique()
                    dados_validos_df = df_merge_mae.dropna(subset=['semi'])

            # ------------------------------------------------------------------
            # 2) TRATAMENTO DE CÓDIGOS QUE NÃO ESTÃO NA PLANILHA MÃE
            # ------------------------------------------------------------------
            if len(codigos_faltantes) > 0:
                st.warning(f"⚠️ {len(codigos_faltantes)} códigos não encontrados na Planilha Mãe!")
                
                col1, col2 = st.columns(2)
                with col1:
                    df_faltantes = pd.DataFrame({'codigo': codigos_faltantes})
                    df_faltantes['semi'] = ''
                    df_faltantes['gola'] = ''
                    df_faltantes['bordado'] = ''
                    df_faltantes['componentes_codigos'] = ''
                    excel_faltantes = gerar_excel_formatado(df_faltantes, "codigos_faltantes")
                    st.download_button(
                        label="📥 Baixar Códigos Faltantes",
                        data=excel_faltantes,
                        file_name="codigos_faltantes.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                with col2:
                    uploaded_faltantes = st.file_uploader(
                        "📤 Enviar Códigos Completados", type=["xlsx"], key="codigos_completados",
                        help="Preencha e envie a planilha de códigos faltantes."
                    )
                    if uploaded_faltantes:
                        try:
                            df_novos = load_excel(uploaded_faltantes)
                            df_novos.columns = df_novos.columns.str.strip().str.replace(" ", "_").str.lower()
                            
                            if 'codigo' in df_novos.columns:
                                df_mae_atualizada = pd.concat([df_mae, df_novos], ignore_index=True)
                                df_mae_atualizada = df_mae_atualizada.drop_duplicates(subset=['codigo'], keep='last')
                                
                                st.session_state['df_mae'] = df_mae_atualizada
                                st.success(f"✅ {len(df_novos)} produtos adicionados/atualizados na Planilha Mãe da sessão!")
                                st.info("🔄 A página será recarregada para aplicar as mudanças. Por favor, reenvie o arquivo de vendas.")
                                
                                excel_mae_atualizada = gerar_excel_formatado(df_mae_atualizada, "planilha_mae_atualizada")
                                st.download_button(
                                    label="📥 Baixar Planilha Mãe Atualizada",
                                    data=excel_mae_atualizada,
                                    file_name="planilha_mae_atualizada.xlsx"
                                )
                                st.rerun()
                            else:
                                st.error("❌ Planilha de códigos completados deve ter a coluna 'codigo'.")
                        except Exception as e:
                            st.error(f"Erro ao processar códigos completados: {str(e)}")

            # ------------------------------------------------------------------
            # 3) EXPLOSÃO DOS PRODUTOS EM INSUMOS (APENAS FALTANTES)
            # ------------------------------------------------------------------
            if not dados_validos_df.empty:
                with st.spinner("Explodindo produtos faltantes em Semi / Gola / Bordado..."):
                    dados_explodidos = explodir_kits(dados_validos_df, st.session_state['df_mae'])

                st.success(f"✅ Explosão concluída! {len(dados_explodidos)} componentes individuais (Semi/Gola/Bordado) encontrados.")

                # ------------------------------------------------------------------
                # 📈 Resumo do Dia (insumos)
                # ------------------------------------------------------------------
                st.header("📈 Resumo do Dia (Componentes para Produzir)")

                col1, col2, col3 = st.columns(3)
                resumos = {
                    "👔 Manga Longa": dados_explodidos[dados_explodidos['semi'].str.contains('Manga Longa', na=False)],
                    "👗 Manga Curta": dados_explodidos[dados_explodidos['semi'].str.contains('Manga Curta', na=False)],
                    "👶 Mijões": dados_explodidos[dados_explodidos['semi'].str.contains('Mijão|Mijao', na=False, regex=True)]
                }
                
                for i, (titulo, df_resumo) in enumerate(resumos.items()):
                    with [col1, col2, col3][i]:
                        st.subheader(titulo)
                        total = df_resumo['quantidade'].sum()
                        if total > 0:
                            st.metric(f"Total {titulo.split(' ')[1]}", int(total))
                        else:
                            st.info(f"Nenhuma necessidade de {titulo.split(' ')[1]} hoje (pela explosão).")

                # ------------------------------------------------------------------
                # 📊 Relatórios “clássicos” (mantidos)
                # ------------------------------------------------------------------
                st.subheader("📊 Relatórios de Componentes (mantidos)")

                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    excel_componentes = gerar_excel_formatado(
                        dados_explodidos[['semi', 'gola', 'bordado', 'quantidade']].copy(),
                        "relatorio_componentes",
                        agrupar_por_semi=True
                    )
                    st.download_button("📋 Componentes (Semi + Golas/Bordados)",
                                       excel_componentes,
                                       "relatorio_componentes.xlsx")

                with col2:
                    resumo_semis = dados_explodidos.groupby('semi', as_index=False)['quantidade'].sum()
                    resumo_semis[['cat', 'cor', 'tam']] = resumo_semis['semi'].apply(
                        lambda x: pd.Series(get_categoria_ordem(x))
                    )
                    resumo_semis = resumo_semis.sort_values(['cat', 'cor', 'tam', 'semi']).drop(columns=['cat', 'cor', 'tam'])
                    resumo_semis = resumo_semis.rename(columns={'semi': 'Semi', 'quantidade': 'Quantidade'})
                    excel_semis = gerar_excel_formatado(resumo_semis, "resumo_semis")
                    st.download_button("📊 Resumo Semis", excel_semis, "resumo_semis.xlsx")
                
                with col3:
                    relatorio_golas = dados_explodidos.groupby('gola')['quantidade'].sum().reset_index()
                    relatorio_golas = relatorio_golas[relatorio_golas['gola'].astype(str).str.strip() != '']
                    relatorio_golas = relatorio_golas.sort_values('quantidade', ascending=False)
                    excel_golas = gerar_excel_formatado(relatorio_golas, "relatorio_golas")
                    st.download_button("👔 Relatório Golas", excel_golas, "relatorio_golas.xlsx")
                
                with col4:
                    relatorio_bordados = dados_explodidos.groupby('bordado')['quantidade'].sum().reset_index()
                    relatorio_bordados = relatorio_bordados[relatorio_bordados['bordado'].astype(str).str.strip() != '']
                    relatorio_bordados = relatorio_bordados.sort_values('quantidade', ascending=False)
                    excel_bordados = gerar_excel_formatado(relatorio_bordados, "relatorio_bordados")
                    st.download_button("🎨 Relatório Bordados", excel_bordados, "relatorio_bordados.xlsx")

                # ------------------------------------------------------------------
                # 💣 NOVO: PRODUZIR HOJE (com base no estoque de insumos)
                # ------------------------------------------------------------------
                st.header("🧱 Produção Necessária Hoje (cruzando com estoque de insumos)")

                if df_estoque_raw is None:
                    st.warning("⚠️ Estoque `template_estoque` não carregado. Não é possível calcular falta de Semi/Gola/Bordado.")
                else:
                    semis_falta, golas_falta, bords_falta = cruzar_com_estoque_insumos(dados_explodidos, df_estoque_raw)

                    # --- SEMIS FALTANTES ---
                    st.subheader("1️⃣ Produzir Hoje – SEMIS")
                    if semis_falta is None or semis_falta.empty:
                        st.success("Nenhum Semi faltando (considerando estoque atual).")
                    else:
                        st.dataframe(semis_falta, use_container_width=True)
                        excel_semis_falta = gerar_excel_formatado(semis_falta, "produzir_semis_hoje")
                        st.download_button(
                            "📥 Baixar 'Produzir Hoje – Semis'",
                            excel_semis_falta,
                            "produzir_semis_hoje.xlsx"
                        )

                    # --- GOLAS FALTANTES ---
                    st.subheader("2️⃣ Produzir Hoje – GOLAS")
                    if golas_falta is None or golas_falta.empty:
                        st.success("Nenhuma Gola faltando (considerando estoque atual).")
                    else:
                        st.dataframe(golas_falta, use_container_width=True)
                        excel_golas_falta = gerar_excel_formatado(golas_falta, "produzir_golas_hoje")
                        st.download_button(
                            "📥 Baixar 'Produzir Hoje – Golas'",
                            excel_golas_falta,
                            "produzir_golas_hoje.xlsx"
                        )

                    # --- BORDAR HOJE (BORDADOS) ---
                    st.subheader("3️⃣ Bordar Hoje – BORDADOS (quando não há gola pronta)")
                    if bords_falta is None or bords_falta.empty:
                        st.success("Nenhum Bordado faltando (considerando estoque atual).")
                    else:
                        st.dataframe(bords_falta, use_container_width=True)
                        excel_bords_falta = gerar_excel_formatado(bords_falta, "bordar_hoje_bordados")
                        st.download_button(
                            "📥 Baixar 'Bordar Hoje – Bordados'",
                            excel_bords_falta,
                            "bordar_hoje_bordados.xlsx"
                        )

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante o processamento: {str(e)}")

# --- Barra Lateral ---
st.sidebar.markdown("---")
st.sidebar.info(
    "💡 **Planilha Mãe:**\n"
    "Fica carregada durante toda esta sessão. "
    "Se fechar o navegador, precisa carregar de novo."
)
st.sidebar.markdown("---")
st.sidebar.info(
    "📦 **Estoque (`template_estoque`):**\n"
    "Usado apenas para LEITURA.\n\n"
    "1️⃣ Primeiro verifica se existe produto pronto.\n"
    "2️⃣ Explode apenas o que está faltando.\n"
    "3️⃣ Depois cruza Semi / Gola / Bordado com o estoque para montar:\n"
    "   • Produzir Hoje – Semis\n"
    "   • Produzir Hoje – Golas\n"
    "   • Bordar Hoje – Bordados"
)
