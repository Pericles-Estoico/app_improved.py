import streamlit as st
import pandas as pd

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
                    
                    # Mostrar códigos faltantes
                    with st.expander("Ver códigos faltantes"):
                        st.write(list(codigos_faltantes))
                
                if not dados_validos.empty:
                    st.success(f"✅ Processando {len(dados_validos)} itens válidos")
                    
                    # Resumo do Dia
                    st.header("📈 Resumo do Dia")
                    
                    col1, col2, col3 = st.columns(3)
                    
                    with col1:
                        st.subheader("👔 Manga Longa")
                        ml_resumo = dados_validos[dados_validos['semi'].str.contains('Manga Longa', na=False)]
                        if not ml_resumo.empty:
                            total_ml = ml_resumo['quantidade'].sum()
                            st.metric("Total ML", total_ml)
                        else:
                            st.info("Nenhuma venda ML hoje")
                    
                    with col2:
                        st.subheader("👗 Manga Curta")
                        mc_resumo = dados_validos[dados_validos['semi'].str.contains('Manga Curta', na=False)]
                        if not mc_resumo.empty:
                            total_mc = mc_resumo['quantidade'].sum()
                            st.metric("Total MC", total_mc)
                        else:
                            st.info("Nenhuma venda MC hoje")
                    
                    with col3:
                        st.subheader("👶 Mijões")
                        mij_resumo = dados_validos[dados_validos['semi'].str.contains('Mijão|Mijao', na=False)]
                        if not mij_resumo.empty:
                            total_mij = mij_resumo['quantidade'].sum()
                            st.metric("Total Mijões", total_mij)
                        else:
                            st.info("Nenhuma venda Mijão hoje")
                    
                    # Mostrar dados processados
                    st.subheader("📋 Dados Processados")
                    st.dataframe(dados_validos[['código', 'quantidade', 'semi', 'gola', 'bordado']], use_container_width=True)
                    
            else:
                st.error("Planilha deve ter colunas 'código' e 'quantidade'")
                
        except Exception as e:
            st.error(f"Erro ao processar vendas: {str(e)}")

# Controle de Estoque Simples
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
st.markdown("**Pure & Posh Baby** - Sistema de Relatórios v1.0 (Cloud)")
