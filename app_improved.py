import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import numpy as np # Adicionado para lidar com valores nulos (NaN)

# ==============================================================================
# CONFIGURAÇÕES E ESTILOS
# ==============================================================================

st.set_page_config(page_title="Pure & Posh Baby - Sistema de Relatórios", page_icon="👑", layout="wide")

# Header
st.markdown("""
<style>
.centered-title { text-align: center; width: 100%; margin: 0 auto; }
@media (max-width: 768px) { .centered-title { text-align: center; } }
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios de Vendas v2.0")
st.markdown("**Pure & Posh Baby**")
st.markdown('</div>', unsafe_allow_html=True)

# Inicializar session_state
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

def get_categoria_ordem(semi):
    """Determina a categoria e a ordem de um item 'semi' para ordenação nos relatórios."""
    semi_str = str(semi).lower()
    
    # Mapeamentos para clareza e facilidade de manutenção
    CATEGORIAS = {
        'manga longa': 1,
        'manga curta menina': 2,
        'manga curta menino': 3,
        'mijão': 4,
        'mijao': 4
    }
    CORES = {'branco': 1, 'vermelho': 2, 'marinho': 3, 'azul': 4, 'rosa': 5}
    TAMANHOS = {'-rn': 1, '-p': 2, '-m': 3, '-g': 4}

    categoria = next((cat for key, cat in CATEGORIAS.items() if key in semi_str), 5)
    cor_ordem = next((cor for key, cor in CORES.items() if key in semi_str), 6)
    tamanho_ordem = next((tam for key, tam in TAMANHOS.items() if key in semi_str), 5)
    
    return categoria, cor_ordem, tamanho_ordem

def explodir_kits(df_vendas_com_mae, df_mae_completa):
    """
    Função principal para "explodir" kits em seus componentes individuais.
    Esta é a nova lógica central do sistema.
    """
    componentes_finais = []
    
    # Garante que o índice do df_mae seja a coluna 'codigo' para buscas rápidas
    df_mae_completa = df_mae_completa.set_index('codigo')

    def obter_componentes(codigo, quantidade):
        """Função recursiva interna para encontrar todos os componentes de um código."""
        lista_componentes_recursiva = []
        
        try:
            produto = df_mae_completa.loc[codigo]
        except KeyError:
            # Se o código não for encontrado, retorna uma lista vazia.
            # O tratamento de códigos faltantes já acontece antes.
            return []

        # 1. Adiciona componentes diretos do produto (se existirem)
        # CORREÇÃO: Verificação mais robusta para evitar erro de Series ambiguous
        semi_valido = False
        if 'semi' in produto.index:
            if pd.notna(produto['semi']):
                if isinstance(produto['semi'], str) and produto['semi'].strip() != '':
                    semi_valido = True

        if semi_valido:
            lista_componentes_recursiva.append({
                'semi': produto['semi'],
                'gola': produto['gola'] if pd.notna(produto['gola']) else '',
                'bordado': produto['bordado'] if pd.notna(produto['bordado']) else '',
                'quantidade': quantidade
            })

        # 2. Processa componentes aninhados (se existirem)
        # CORREÇÃO: Verificação mais robusta para componentes_codigos
        componentes_codigos_valido = False
        if 'componentes_codigos' in produto.index:
            if pd.notna(produto['componentes_codigos']):
                componentes_str = str(produto['componentes_codigos']).strip()
                if componentes_str != '' and componentes_str.lower() != 'nan':
                    componentes_codigos_valido = True

        if componentes_codigos_valido:
            codigos_aninhados = str(produto['componentes_codigos']).split(';')
            for cod_aninhado in codigos_aninhados:
                cod_aninhado = cod_aninhado.strip()
                if cod_aninhado:
                    # Chamada recursiva para explodir os sub-componentes
                    lista_componentes_recursiva.extend(obter_componentes(cod_aninhado, quantidade))
        
        return lista_componentes_recursiva

    # Itera sobre cada linha da planilha de vendas mesclada
    for _, venda in df_vendas_com_mae.iterrows():
        componentes_finais.extend(obter_componentes(venda['codigo'], venda['quantidade']))

    return pd.DataFrame(componentes_finais)


def gerar_excel_formatado(df, nome_arquivo, agrupar_por_semi=False):
    """Gera um arquivo Excel formatado a partir de um DataFrame."""
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
    border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
    
    if agrupar_por_semi:
        headers = ['Item', 'Quantidade', 'Check']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
        
        # Limpar valores nulos antes de agrupar
        df['gola'] = df['gola'].fillna('')
        df['bordado'] = df['bordado'].fillna('')
        
        relatorio_componentes = df.groupby(['semi', 'gola', 'bordado'])['quantidade'].sum().reset_index()
        
        relatorio_componentes[['categoria', 'cor_ordem', 'tamanho_ordem']] = relatorio_componentes['semi'].apply(
            lambda x: pd.Series(get_categoria_ordem(x))
        )
        
        relatorio_componentes = relatorio_componentes.sort_values(
            ['categoria', 'cor_ordem', 'tamanho_ordem', 'semi', 'gola', 'bordado']
        )
        
        relatorio_hierarquico = []
        for semi_produto, grupo in relatorio_componentes.groupby('semi'):
            total_semi = grupo['quantidade'].sum()
            categoria = grupo['categoria'].iloc[0]
            
            relatorio_hierarquico.append({
                'Item': semi_produto, 'Quantidade': total_semi, 'Check': '', 'categoria': categoria, 'is_semi': True
            })
            
            for _, row in grupo.iterrows():
                componente = f"{row['gola']} {row['bordado']}".strip()
                if componente: # Só adiciona se houver gola ou bordado
                    relatorio_hierarquico.append({
                        'Item': f"  {componente}", 'Quantidade': row['quantidade'], 'Check': '', 'categoria': categoria, 'is_semi': False
                    })

        row_num = 2
        for item in relatorio_hierarquico:
            is_semi = item['is_semi']
            categoria = item['categoria']
            
            fill_color = None
            if is_semi:
                if categoria == 1: fill_color = manga_longa_fill
                elif categoria == 2: fill_color = manga_curta_menina_fill
                elif categoria == 3: fill_color = manga_curta_menino_fill
                elif categoria == 4: fill_color = mijao_fill
            
            for col_num, key in enumerate(['Item', 'Quantidade', 'Check'], 1):
                cell = ws.cell(row=row_num, column=col_num, value=item[key])
                cell.border = border
                if is_semi:
                    if col_num == 1: cell.font = semi_font
                    if fill_color: cell.fill = fill_color
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
        
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except: pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output)
    output.seek(0)
    return output

# ==============================================================================
# INTERFACE DO STREAMLIT
# ==============================================================================

st.header("📁 Configuração Inicial")

def carregar_planilha_mae(arquivo):
    """Lógica para carregar e validar a planilha mãe."""
    try:
        with st.spinner("Carregando e validando Planilha Mãe..."):
            df = load_excel(arquivo)
            df.columns = df.columns.str.strip().str.replace(" ", "_").str.lower()
            
            # Validação das colunas essenciais
            colunas_essenciais = ['codigo', 'semi', 'gola', 'bordado']
            if not all(col in df.columns for col in colunas_essenciais):
                st.error(f"❌ Erro: A Planilha Mãe deve conter as colunas: {', '.join(colunas_essenciais)}.")
                return False

            # Adiciona a coluna 'componentes_codigos' se não existir, para retrocompatibilidade
            if 'componentes_codigos' not in df.columns:
                df['componentes_codigos'] = ''

            st.session_state['df_mae'] = df
            st.session_state['planilha_mae_carregada'] = True
            st.success(f"✅ Planilha Mãe carregada: {len(df)} produtos cadastrados.")
            st.rerun()
    except Exception as e:
        st.error(f"Erro ao carregar planilha mãe: {str(e)}")

if st.session_state['planilha_mae_carregada']:
    st.success(f"✅ Planilha Mãe carregada: {len(st.session_state['df_mae'])} produtos cadastrados.")
    with st.expander("🔄 Recarregar/Atualizar Planilha Mãe"):
        uploaded_mae_nova = st.file_uploader("Substitua a Planilha Mãe atual", type=["xlsx"], key="planilha_mae_nova")
        if uploaded_mae_nova:
            carregar_planilha_mae(uploaded_mae_nova)
else:
    st.info("📋 Para começar, carregue a Planilha Mãe. Ela deve conter as colunas: `codigo`, `semi`, `gola`, `bordado` e, opcionalmente, `componentes_codigos` para kits.")
    uploaded_mae = st.file_uploader("Carregar Planilha Mãe", type=["xlsx"], key="planilha_mae")
    if uploaded_mae:
        carregar_planilha_mae(uploaded_mae)

# --- Processamento de Vendas ---
if st.session_state['planilha_mae_carregada']:
    st.header("📊 Processamento Diário")
    
    uploaded_vendas = st.file_uploader("📈 Planilha de Vendas (diária)", type=["xlsx"], key="vendas")
    
    if uploaded_vendas:
        try:
            with st.spinner("Processando vendas..."):
                df_vendas = load_excel(uploaded_vendas)
                df_vendas.columns = df_vendas.columns.str.strip().str.replace(' ', '_').str.lower()

                if 'código' not in df_vendas.columns or 'quantidade' not in df_vendas.columns:
                    st.error("❌ Planilha de vendas deve ter colunas 'código' e 'quantidade'")
                    st.stop()

                df_vendas = df_vendas.rename(columns={'código': 'codigo'})
                df_mae = st.session_state['df_mae']
                
                # Mescla para encontrar códigos faltantes
                df_merged = pd.merge(df_vendas[['codigo', 'quantidade']], df_mae, on='codigo', how='left')
                
                codigos_faltantes = df_merged[df_merged['semi'].isna()]['codigo'].unique()
                dados_validos_df = df_merged.dropna(subset=['semi'])

            if len(codigos_faltantes) > 0:
                st.warning(f"⚠️ {len(codigos_faltantes)} códigos não encontrados na Planilha Mãe!")
                
                col1, col2 = st.columns(2)
                with col1:
                    df_faltantes = pd.DataFrame({'codigo': codigos_faltantes})
                    # Adiciona colunas para preenchimento
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

            if not dados_validos_df.empty:
                with st.spinner("Explodindo kits e gerando relatórios..."):
                    # AQUI A MÁGICA ACONTECE
                    dados_explodidos = explodir_kits(dados_validos_df, df_mae)

                st.success(f"✅ Processamento concluído! {len(dados_explodidos)} componentes individuais encontrados.")
                
                # Resumo do Dia
                st.header("📈 Resumo do Dia (Componentes)")
                col1, col2, col3 = st.columns(3)
                
                resumos = {
                    "👔 Manga Longa": dados_explodidos[dados_explodidos['semi'].str.contains('Manga Longa', na=False)],
                    "👗 Manga Curta": dados_explodidos[dados_explodidos['semi'].str.contains('Manga Curta', na=False)],
                    "👶 Mijões": dados_explodidos[dados_explodidos['semi'].str.contains('Mijão|Mijao', na=False)]
                }
                
                for i, (titulo, df_resumo) in enumerate(resumos.items()):
                    with [col1, col2, col3][i]:
                        st.subheader(titulo)
                        total = df_resumo['quantidade'].sum()
                        if total > 0:
                            st.metric(f"Total {titulo.split(' ')[1]}", int(total))
                        else:
                            st.info(f"Nenhuma venda de {titulo.split(' ')[1]} hoje.")

                # Relatórios para Download
                st.subheader("📊 Relatórios para Download")
                col1, col2, col3, col4 = st.columns(4)
                
                with col1:
                    excel_componentes = gerar_excel_formatado(dados_explodidos, "relatorio_componentes", agrupar_por_semi=True)
                    st.download_button("📋 Relatório Componentes", excel_componentes, "relatorio_componentes.xlsx")
                
                with col2:
                    resumo_semis = dados_explodidos.groupby('semi')['quantidade'].sum().reset_index()
                    resumo_semis[['cat', 'cor', 'tam']] = resumo_semis['semi'].apply(lambda x: pd.Series(get_categoria_ordem(x)))
                    resumo_semis = resumo_semis.sort_values(['cat', 'cor', 'tam', 'semi']).drop(columns=['cat', 'cor', 'tam'])
                    excel_semis = gerar_excel_formatado(resumo_semis, "resumo_semis")
                    st.download_button("📊 Resumo Semis", excel_semis, "resumo_semis.xlsx")
                
                with col3:
                    relatorio_golas = dados_explodidos.groupby('gola')['quantidade'].sum().reset_index().sort_values('quantidade', ascending=False)
                    excel_golas = gerar_excel_formatado(relatorio_golas, "relatorio_golas")
                    st.download_button("👔 Relatório Golas", excel_golas, "relatorio_golas.xlsx")
                
                with col4:
                    relatorio_bordados = dados_explodidos.groupby('bordado')['quantidade'].sum().reset_index().sort_values('quantidade', ascending=False)
                    excel_bordados = gerar_excel_formatado(relatorio_bordados, "relatorio_bordados")
                    st.download_button("🎨 Relatório Bordados", excel_bordados, "relatorio_bordados.xlsx")

        except Exception as e:
            st.error(f"Ocorreu um erro inesperado durante o processamento: {str(e)}")

# --- Barra Lateral ---
st.sidebar.markdown("---")
st.sidebar.info("💡 **Sobre a Planilha Mãe:**\n\nA planilha fica carregada durante toda esta sessão. Se fechar e abrir o navegador, precisará carregá-la novamente.")
st.sidebar.markdown("---")
st.sidebar.info("📦 **Como Cadastrar Kits:**\n\n1. Crie uma linha para o código do kit.\n2. Na coluna `componentes_codigos`, liste os códigos dos produtos que formam o kit, separados por `;`.\n3. Se o kit também tiver um componente direto (ex: um body), preencha as colunas `semi`, `gola` e `bordado` na mesma linha do kit.")
