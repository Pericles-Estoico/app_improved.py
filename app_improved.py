import streamlit as st
import pandas as pd
from io import BytesIO

# Configurações
st.set_page_config(page_title="Pure & Posh Baby - Sistema de Relatórios", page_icon="👑", layout="wide")

# Header
st.markdown("""
<style>
.centered-title {
    text-align: center;
    width: 100%;
    margin: 0 auto;
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios de Vendas")
st.markdown("**Pure & Posh Baby**")
st.markdown('</div>', unsafe_allow_html=True)

# Função para carregar Excel
def load_excel(arquivo):
    return pd.read_excel(arquivo)

# Função para gerar Excel para download
def gerar_excel_download(df, nome_arquivo):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Relatório')
    output.seek(0)
    return output

# Interface principal
st.header("📁 Configuração Inicial")

# Upload da Planilha Mãe
uploaded_mae = st.file_uploader("📋 Carregar Planilha Mãe", type=["xlsx"], key="planilha_mae")

if uploaded_mae:
    try:
        df_mae = load_excel(uploaded_mae)
        df_mae.columns = df_mae.columns.str.strip().str.replace(" ", "_").str.lower()
        st.success(f"✅ Planilha Mãe carregada: {len(df_mae)} registros")
        
        # Armazenar na sessão
        st.session_state['df_mae'] = df_mae
        
    except Exception as e:
        st.error(f"Erro ao carregar planilha mãe: {str(e)}")

# Processamento de vendas
if 'df_mae' in st.session_state:
    st.header("📊 Processamento Diário")
    
    uploaded_vendas = st.file_uploader("📈 Planilha de Vendas (diária)", type=["xlsx"], key="vendas")
    
    if uploaded_vendas:
        try:
            df_vendas = load_excel(uploaded_vendas)
            df_vendas.columns = df_vendas.columns.str.strip().str.replace(' ', '_').str.lower()
            
            if 'código' in df_vendas.columns and 'quantidade' in df_vendas.columns:
                df_mae = st.session_state['df_mae']
                
                # Mesclar dados
                df_final = pd.merge(df_vendas, df_mae, left_on='código', right_on='codigo', how='left')
                
                # Códigos faltantes
                codigos_faltantes = df_final[df_final['semi'].isna()]['código'].unique()
                dados_validos = df_final[df_final['semi'].notna()].copy()
                
                if len(codigos_faltantes) > 0:
                    st.warning(f"⚠️ {len(codigos_faltantes)} códigos faltantes")
                    
                    # Download códigos faltantes
                    df_faltantes = pd.DataFrame({'codigo': codigos_faltantes})
                    excel_faltantes = gerar_excel_download(df_faltantes, "codigos_faltantes")
                    st.download_button(
                        label="📥 Baixar Códigos Faltantes",
                        data=excel_faltantes,
                        file_name="codigos_faltantes.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                if not dados_validos.empty:
                    st.success(f"✅ Gerando relatórios com {len(dados_validos)} itens")
                    
                    # Resumo do Dia
                    st.header("📈 Resumo do Dia")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.subheader("👔 Manga Longa")
                        ml_resumo = dados_validos[dados_validos['semi'].str.contains('Manga Longa|ML|manga longa', na=False, case=False)]
                        if not ml_resumo.empty:
                            total_ml = ml_resumo['quantidade'].sum()
                            st.metric("Total ML", total_ml)
                        else:
                            st.info("Nenhuma venda ML hoje")
                    
                    with col2:
                        st.subheader("👗 Manga Curta")
                        mc_resumo = dados_validos[dados_validos['semi'].str.contains('Manga Curta|MC|manga curta', na=False, case=False)]
                        if not mc_resumo.empty:
                            total_mc = mc_resumo['quantidade'].sum()
                            st.metric("Total MC", total_mc)
                        else:
                            st.info("Nenhuma venda MC hoje")
                    
                    with col3:
                        st.subheader("👶 Mijões")
                        mij_resumo = dados_validos[dados_validos['semi'].str.contains('Mijão|Mijao|mijão|mijao', na=False, case=False)]
                        if not mij_resumo.empty:
                            total_mij = mij_resumo['quantidade'].sum()
                            st.metric("Total Mijões", total_mij)
                        else:
                            st.info("Nenhuma venda Mijão hoje")
                    
                    # Relatórios para Download
                    st.subheader("📊 Relatórios para Download")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        # Relatório de Componentes
                        relatorio_componentes = dados_validos.groupby(['semi', 'gola', 'bordado'])['quantidade'].sum().reset_index()
                        excel_componentes = gerar_excel_download(relatorio_componentes, "relatorio_componentes")
                        st.download_button(
                            label="📥 Relatório Componentes",
                            data=excel_componentes,
                            file_name="relatorio_componentes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col2:
                        # Resumo Semis
                        resumo_semis = dados_validos.groupby('semi')['quantidade'].sum().reset_index()
                        excel_semis = gerar_excel_download(resumo_semis, "resumo_semis")
                        st.download_button(
                            label="📥 Resumo Semis",
                            data=excel_semis,
                            file_name="resumo_semis.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col3:
                        # Relatório Golas
                        relatorio_golas = dados_validos.groupby('gola')['quantidade'].sum().reset_index()
                        excel_golas = gerar_excel_download(relatorio_golas, "relatorio_golas")
                        st.download_button(
                            label="📥 Relatório Golas",
                            data=excel_golas,
                            file_name="relatorio_golas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col4:
                        # Relatório Bordados
                        relatorio_bordados = dados_validos.groupby('bordado')['quantidade'].sum().reset_index()
                        excel_bordados = gerar_excel_download(relatorio_bordados, "relatorio_bordados")
                        st.download_button(
                            label="📥 Relatório Bordados",
                            data=excel_bordados,
                            file_name="relatorio_bordados.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
            else:
                st.error("Planilha deve ter colunas 'código' e 'quantidade'")
                
        except Exception as e:
            st.error(f"Erro ao processar vendas: {str(e)}")

# Controle de Estoque
st.header("📦 Controle de Estoque")

if 'df_mae' in st.session_state:
    df_mae = st.session_state['df_mae']
    
    if 'codigo' in df_mae.columns:
        produtos_lista = df_mae['codigo'].tolist()
        
        # Busca de produto
        selected_items = st.multiselect(
            "Busque e selecione o item:",
            options=produtos_lista,
            max_selections=1,
            placeholder="Digite para buscar..."
        )
        
        if selected_items:
            selected_item = selected_items[0]
            
            # Mostrar informações do produto
            produto_info = df_mae[df_mae['codigo'] == selected_item].iloc[0]
            
            col1, col2 = st.columns(2)
            with col1:
                st.info(f"**Código:** {selected_item}")
                st.info(f"**Semi:** {produto_info.get('semi', 'N/A')}")
            with col2:
                st.info(f"**Gola:** {produto_info.get('gola', 'N/A')}")
                st.info(f"**Bordado:** {produto_info.get('bordado', 'N/A')}")
    else:
        st.warning("Planilha mãe deve ter coluna 'codigo'")
else:
    st.info("Carregue a Planilha Mãe primeiro para usar o controle de estoque")

st.markdown("---")
st.markdown("**Pure & Posh Baby** - Sistema de Relatórios v1.0")
