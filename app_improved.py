import streamlit as st
import pandas as pd
from io import BytesIO
import os
import pickle
import datetime

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

# Arquivos
PLANILHA_MAE_FILE = "planilha_mae_fixa.pkl"
ESTOQUE_FILE = "estoque.pkl"

# Funções
def load_excel(arquivo):
    return pd.read_excel(arquivo)

def salvar_planilha_mae(df):
    with open(PLANILHA_MAE_FILE, 'wb') as f:
        pickle.dump(df, f)

def carregar_planilha_mae():
    if os.path.exists(PLANILHA_MAE_FILE):
        with open(PLANILHA_MAE_FILE, 'rb') as f:
            return pickle.load(f)
    return None

def salvar_estoque(estoque_df):
    with open(ESTOQUE_FILE, 'wb') as f:
        pickle.dump(estoque_df, f)

def carregar_estoque():
    if os.path.exists(ESTOQUE_FILE):
        with open(ESTOQUE_FILE, 'rb') as f:
            return pickle.load(f)
    return pd.DataFrame(columns=['codigo', 'semi', 'gola', 'bordado', 'quantidade'])

# Interface principal
st.header("📁 Configuração Inicial")

# Planilha mãe
df_mae = carregar_planilha_mae()
if df_mae is not None:
    st.success(f"✅ Planilha Mãe compartilhada: {len(df_mae)} registros")
else:
    st.warning("⚠️ Planilha Mãe não configurada")
    uploaded_mae = st.file_uploader("📋 Carregar Planilha Mãe (uma vez)", type=["xlsx"])
    if uploaded_mae:
        try:
            df_mae = load_excel(uploaded_mae)
            df_mae.columns = df_mae.columns.str.strip().str.replace(" ", "_").str.lower()
            salvar_planilha_mae(df_mae)
            st.success("✅ Planilha Mãe salva no sistema!")
            st.rerun()
        except Exception as e:
            st.error(f"Erro ao carregar planilha: {str(e)}")

# Processamento diário
if df_mae is not None:
    st.header("📊 Processamento Diário")
    
    uploaded_vendas = st.file_uploader("📈 Planilha de Vendas (diária)", type=["xlsx"])
    
    if uploaded_vendas:
        try:
            df_vendas = load_excel(uploaded_vendas)
            df_vendas.columns = df_vendas.columns.str.strip().str.replace(' ', '_').str.lower()
            
            if 'código' in df_vendas.columns and 'quantidade' in df_vendas.columns:
                # Mesclar
                df_final = pd.merge(df_vendas, df_mae, left_on='código', right_on='codigo', how='left')
                
                # Códigos faltantes
                codigos_faltantes = df_final[df_final['semi'].isna()]['código'].unique()
                dados_validos = df_final[df_final['semi'].notna()].copy()
                
                if len(codigos_faltantes) > 0:
                    st.warning(f"⚠️ {len(codigos_faltantes)} códigos faltantes")
                
                if not dados_validos.empty:
                    st.success(f"✅ Gerando relatórios com {len(dados_validos)} itens")
                    
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
            else:
                st.error("Planilha deve ter colunas 'código' e 'quantidade'")
        except Exception as e:
            st.error(f"Erro ao processar vendas: {str(e)}")

# Controle de Estoque
st.header("📦 Controle de Estoque")

if df_mae is not None:
    produtos_lista = df_mae['codigo'].tolist() if 'codigo' in df_mae.columns else []
    
    if produtos_lista:
        selected_items = st.multiselect(
            "Busque e selecione o item:",
            options=produtos_lista,
            max_selections=1,
            placeholder="Digite para buscar..."
        )
        
        if selected_items:
            selected_item = selected_items[0]
            quantidade = st.number_input("Quantidade a Adicionar/Remover", value=0, step=1)
            
            if st.button("Adicionar/Atualizar Estoque"):
                if quantidade != 0:
                    try:
                        produto_info = df_mae[df_mae['codigo'] == selected_item].iloc[0]
                        estoque_df = carregar_estoque()
                        
                        idx = estoque_df[estoque_df['codigo'] == selected_item].index
                        
                        if not idx.empty:
                            estoque_df.loc[idx, 'quantidade'] += quantidade
                        else:
                            novo_item = pd.DataFrame([{
                                'codigo': selected_item,
                                'semi': produto_info.get('semi', ''),
                                'gola': produto_info.get('gola', ''),
                                'bordado': produto_info.get('bordado', ''),
                                'quantidade': quantidade
                            }])
                            estoque_df = pd.concat([estoque_df, novo_item], ignore_index=True)
                        
                        salvar_estoque(estoque_df)
                        st.success(f"✅ Estoque atualizado! {selected_item}: {quantidade:+d}")
                    except Exception as e:
                        st.error(f"Erro ao atualizar estoque: {str(e)}")
                else:
                    st.warning("Digite uma quantidade diferente de zero")
    else:
        st.info("Carregue a Planilha Mãe primeiro")

# Estoque atual
st.subheader("Estoque Atual - Resumo")
try:
    estoque_atual = carregar_estoque()
    
    if not estoque_atual.empty:
        estoque_positivo = estoque_atual[estoque_atual['quantidade'] > 0]
        
        if not estoque_positivo.empty:
            st.dataframe(estoque_positivo[['semi', 'gola', 'bordado', 'quantidade']], use_container_width=True)
            st.info(f"Total de itens em estoque: {len(estoque_positivo)}")
        else:
            st.info("Estoque vazio")
    else:
        st.info("Nenhum item no estoque")
except Exception as e:
    st.error(f"Erro ao carregar estoque: {str(e)}")

st.markdown("---")
st.markdown("**Pure & Posh Baby** - Sistema de Relatórios v1.0")
