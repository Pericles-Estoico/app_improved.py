import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import os
import pickle

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
@media (max-width: 768px) {
    .centered-title {
        text-align: center;
    }
}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="centered-title">', unsafe_allow_html=True)
st.title("👑 Sistema de Relatórios de Vendas")
st.markdown("**Pure & Posh Baby**")
st.markdown('</div>', unsafe_allow_html=True)

# Arquivos para persistência
ESTOQUE_FILE = "estoque.pkl"

# Função para carregar Excel
def load_excel(arquivo):
    return pd.read_excel(arquivo)

# Funções de estoque
def salvar_estoque(estoque_df):
    with open(ESTOQUE_FILE, 'wb') as f:
        pickle.dump(estoque_df, f)

def carregar_estoque():
    if os.path.exists(ESTOQUE_FILE):
        with open(ESTOQUE_FILE, 'rb') as f:
            return pickle.load(f)
    return pd.DataFrame(columns=['codigo', 'semi', 'gola', 'bordado', 'quantidade'])

# Função para determinar categoria e ordem
def get_categoria_ordem(semi):
    semi_str = str(semi).lower()
    
    # Determinar categoria principal
    if 'manga longa' in semi_str:
        categoria = 1  # Azul - primeiro
    elif 'manga curta menina' in semi_str:
        categoria = 2  # Rosa - segundo
    elif 'manga curta menino' in semi_str:
        categoria = 3  # Marinho - terceiro
    elif 'mijão' in semi_str or 'mijao' in semi_str:
        categoria = 4  # Amarelo - quarto
    else:
        categoria = 5  # Outros
    
    # Determinar cor (branco primeiro)
    if 'branco' in semi_str:
        cor_ordem = 1
    elif 'vermelho' in semi_str:
        cor_ordem = 2
    elif 'marinho' in semi_str:
        cor_ordem = 3
    elif 'azul' in semi_str:
        cor_ordem = 4
    elif 'rosa' in semi_str:
        cor_ordem = 5
    else:
        cor_ordem = 6
    
    # Determinar tamanho (RN, P, M, G)
    if '-rn' in semi_str:
        tamanho_ordem = 1
    elif '-p' in semi_str:
        tamanho_ordem = 2
    elif '-m' in semi_str:
        tamanho_ordem = 3
    elif '-g' in semi_str:
        tamanho_ordem = 4
    else:
        tamanho_ordem = 5
    
    return categoria, cor_ordem, tamanho_ordem

# Função para gerar Excel formatado com ordenação correta
def gerar_excel_formatado(df, nome_arquivo, agrupar_por_semi=False):
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Relatório"
    
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True)
    
    # Cores específicas por tipo de produto
    manga_longa_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")  # Azul claro
    manga_curta_menina_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")  # Rosa claro
    manga_curta_menino_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")  # Azul escuro
    mijao_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # Amarelo
    
    semi_font = Font(bold=True)
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    if agrupar_por_semi:
        # Cabeçalhos fixos para relatório de componentes
        headers = ['Item', 'Quantidade', 'Check']
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
        
        # Agrupar dados
        relatorio_componentes = df.groupby(['semi', 'gola', 'bordado'])['quantidade'].sum().reset_index()
        
        # Adicionar colunas de ordenação
        relatorio_componentes[['categoria', 'cor_ordem', 'tamanho_ordem']] = relatorio_componentes['semi'].apply(
            lambda x: pd.Series(get_categoria_ordem(x))
        )
        
        # Ordenar conforme especificado
        relatorio_componentes = relatorio_componentes.sort_values([
            'categoria',      # 1=Manga Longa, 2=MC Menina, 3=MC Menino, 4=Mijão
            'cor_ordem',      # 1=Branco primeiro
            'tamanho_ordem',  # 1=RN, 2=P, 3=M, 4=G
            'semi',
            'gola',
            'bordado'
        ])
        
        # Criar estrutura hierárquica ordenada
        relatorio_hierarquico = []
        current_semi = None
        
        for _, row in relatorio_componentes.iterrows():
            if row['semi'] != current_semi:
                # Adicionar linha do semi
                current_semi = row['semi']
                total_semi = relatorio_componentes[relatorio_componentes['semi'] == current_semi]['quantidade'].sum()
                relatorio_hierarquico.append({
                    'Item': current_semi,
                    'Quantidade': total_semi,
                    'Check': '',
                    'categoria': row['categoria']
                })
            
            # Adicionar linha do componente
            componente = f"{row['gola']} {row['bordado']}".strip()
            relatorio_hierarquico.append({
                'Item': f"  {componente}",
                'Quantidade': row['quantidade'],
                'Check': '',
                'categoria': row['categoria']
            })
        
        # Escrever dados no Excel - APENAS 3 COLUNAS
        row_num = 2
        for item in relatorio_hierarquico:
            item_name = item['Item']
            quantidade = item['Quantidade']
            check = item['Check']
            categoria = item['categoria']
            
            # Determinar cor de fundo
            if not item_name.startswith('  '):  # É um semi
                if categoria == 1:  # Manga Longa
                    semi_fill = manga_longa_fill
                elif categoria == 2:  # MC Menina
                    semi_fill = manga_curta_menina_fill
                elif categoria == 3:  # MC Menino
                    semi_fill = manga_curta_menino_fill
                elif categoria == 4:  # Mijão
                    semi_fill = mijao_fill
                else:
                    semi_fill = manga_longa_fill
                
                # Linha do semi com formatação
                cell1 = ws.cell(row=row_num, column=1, value=item_name)
                cell1.fill = semi_fill
                cell1.font = semi_font
                cell1.border = border
                
                cell2 = ws.cell(row=row_num, column=2, value=quantidade)
                cell2.border = border
                
                cell3 = ws.cell(row=row_num, column=3, value=check)
                cell3.border = border
            else:
                # Linha de componente
                cell1 = ws.cell(row=row_num, column=1, value=item_name)
                cell1.border = border
                
                cell2 = ws.cell(row=row_num, column=2, value=quantidade)
                cell2.border = border
                
                cell3 = ws.cell(row=row_num, column=3, value=check)
                cell3.border = border
            
            row_num += 1
        
        # Ajustar largura das colunas
        ws.column_dimensions['A'].width = 60
        ws.column_dimensions['B'].width = 12
        ws.column_dimensions['C'].width = 8
        
    else:
        # Formato simples - usar colunas do dataframe
        headers = list(df.columns)
        for col_num, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_num, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
        
        for row_num, (_, row) in enumerate(df.iterrows(), 2):
            for col_num, value in enumerate(row, 1):
                cell = ws.cell(row=row_num, column=col_num, value=value)
                cell.border = border
        
        # Ajustar largura das colunas
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
    
    wb.save(output)
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
                    
                    col1, col2 = st.columns(2)
                    
                    with col1:
                        # Download códigos faltantes
                        df_faltantes = pd.DataFrame({'codigo': codigos_faltantes})
                        excel_faltantes = gerar_excel_formatado(df_faltantes, "codigos_faltantes")
                        st.download_button(
                            label="📥 Baixar Códigos Faltantes",
                            data=excel_faltantes,
                            file_name="codigos_faltantes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col2:
                        # Upload códigos faltantes completados
                        uploaded_faltantes = st.file_uploader(
                            "📤 Enviar Códigos Completados", 
                            type=["xlsx"], 
                            key="codigos_completados",
                            help="Envie a planilha de códigos faltantes preenchida com semi, gola e bordado"
                        )
                        
                        if uploaded_faltantes:
                            try:
                                df_novos = load_excel(uploaded_faltantes)
                                df_novos.columns = df_novos.columns.str.strip().str.replace(" ", "_").str.lower()
                                
                                # Verificar se tem as colunas necessárias
                                if all(col in df_novos.columns for col in ['codigo', 'semi', 'gola', 'bordado']):
                                    # Adicionar novos produtos à planilha mãe
                                    df_mae_atualizada = pd.concat([st.session_state['df_mae'], df_novos], ignore_index=True)
                                    df_mae_atualizada = df_mae_atualizada.drop_duplicates(subset=['codigo'], keep='last')
                                    
                                    # Atualizar na sessão
                                    st.session_state['df_mae'] = df_mae_atualizada
                                    
                                    st.success(f"✅ {len(df_novos)} produtos adicionados à planilha mãe!")
                                    st.info("🔄 Reprocesse a planilha de vendas para ver os novos produtos")
                                    
                                    # Botão para baixar planilha mãe atualizada
                                    excel_mae_atualizada = gerar_excel_formatado(df_mae_atualizada, "planilha_mae_atualizada")
                                    st.download_button(
                                        label="📥 Baixar Planilha Mãe Atualizada",
                                        data=excel_mae_atualizada,
                                        file_name="planilha_mae_atualizada.xlsx",
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                    )
                                else:
                                    st.error("❌ Planilha deve ter colunas: codigo, semi, gola, bordado")
                            except Exception as e:
                                st.error(f"Erro ao processar códigos completados: {str(e)}")
                
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
                    
                    # Relatórios para Download
                    st.subheader("📊 Relatórios para Download")
                    
                    col1, col2, col3, col4 = st.columns(4)
                    
                    with col1:
                        # Relatório de Componentes com ordenação correta
                        excel_componentes = gerar_excel_formatado(dados_validos, "relatorio_componentes", agrupar_por_semi=True)
                        
                        st.download_button(
                            label="📥 Relatório Componentes",
                            data=excel_componentes,
                            file_name="relatorio_componentes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col2:
                        # Resumo Semis
                        resumo_semis = dados_validos.groupby('semi')['quantidade'].sum().reset_index()
                        excel_semis = gerar_excel_formatado(resumo_semis, "resumo_semis")
                        st.download_button(
                            label="📥 Resumo Semis",
                            data=excel_semis,
                            file_name="resumo_semis.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col3:
                        # Relatório Golas
                        relatorio_golas = dados_validos.groupby('gola')['quantidade'].sum().reset_index()
                        excel_golas = gerar_excel_formatado(relatorio_golas, "relatorio_golas")
                        st.download_button(
                            label="📥 Relatório Golas",
                            data=excel_golas,
                            file_name="relatorio_golas.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                    with col4:
                        # Relatório Bordados
                        relatorio_bordados = dados_validos.groupby('bordado')['quantidade'].sum().reset_index()
                        excel_bordados = gerar_excel_formatado(relatorio_bordados, "relatorio_bordados")
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

# CONTROLE DE ESTOQUE - CORRIGIDO
if 'df_mae' in st.session_state:
    st.header("📦 Controle de Estoque")
    st.subheader("Adicionar/Atualizar Estoque Manualmente")

    # BUSCA MELHORADA - SEPARANDO TIPOS
    st.write("**Selecione o tipo de entrada:**")
    
    tipo_entrada = st.radio(
        "Tipo de entrada:",
        ["Semi (entrada geral)", "Produto específico (código)"],
        help="Semi = entrada de estoque geral | Produto = entrada de item específico com gola e bordado"
    )
    
    df_mae = st.session_state['df_mae']
    
    if tipo_entrada == "Semi (entrada geral)":
        # Lista apenas os semis únicos
        semis_unicos = sorted(df_mae['semi'].dropna().unique().tolist())
        
        selected_semi = st.selectbox(
            "Selecione o Semi:",
            options=[""] + semis_unicos,
            help="Entrada de estoque geral para o semi (sem gola/bordado específicos)"
        )
        
        if selected_semi:
            st.success(f"Semi selecionado: {selected_semi}")
            
        quantidade_adicionar = st.number_input("Quantidade a Adicionar/Remover", value=0, step=1)

        if st.button("Adicionar/Atualizar Estoque"):
            if selected_semi:
                estoque_df = carregar_estoque()
                
                # Para entrada de semi, usar campos genéricos
                item_semi = selected_semi
                item_gola = "GERAL"  # Marcador para entrada geral
                item_bordado = "GERAL"  # Marcador para entrada geral
                
                # Verificar se já existe no estoque
                existing_idx = estoque_df[
                    (estoque_df['semi'] == item_semi) & 
                    (estoque_df['gola'] == item_gola) & 
                    (estoque_df['bordado'] == item_bordado)
                ].index

                if not existing_idx.empty:
                    estoque_df.loc[existing_idx, 'quantidade'] += quantidade_adicionar
                else:
                    novo_item_estoque = pd.DataFrame([{
                        'codigo': f"SEMI_{selected_semi}", 
                        'semi': item_semi, 
                        'gola': item_gola, 
                        'bordado': item_bordado, 
                        'quantidade': quantidade_adicionar
                    }])
                    estoque_df = pd.concat([estoque_df, novo_item_estoque], ignore_index=True)
                
                salvar_estoque(estoque_df)
                st.success(f"Estoque do semi '{selected_semi}' atualizado!")
            else:
                st.warning("Por favor, selecione um semi.")
    
    else:  # Produto específico
        # Lista todos os códigos
        codigos_unicos = sorted(df_mae['codigo'].dropna().unique().tolist())
        
        selected_codigo = st.selectbox(
            "Selecione o Código do Produto:",
            options=[""] + codigos_unicos,
            help="Entrada de estoque para produto específico (com gola e bordado)"
        )
        
        if selected_codigo:
            # Mostrar detalhes do produto
            produto_info = df_mae[df_mae['codigo'] == selected_codigo].iloc[0]
            st.info(f"**Semi:** {produto_info.get('semi', 'N/A')}")
            st.info(f"**Gola:** {produto_info.get('gola', 'N/A')}")
            st.info(f"**Bordado:** {produto_info.get('bordado', 'N/A')}")
            
        quantidade_adicionar = st.number_input("Quantidade a Adicionar/Remover", value=0, step=1, key="produto_qtd")

        if st.button("Adicionar/Atualizar Estoque", key="produto_btn"):
            if selected_codigo:
                estoque_df = carregar_estoque()
                
                # Pegar informações do produto
                produto_info = df_mae[df_mae['codigo'] == selected_codigo].iloc[0]
                item_semi = produto_info.get('semi', '')
                item_gola = produto_info.get('gola', '')
                item_bordado = produto_info.get('bordado', '')
                
                # Verificar se já existe no estoque
                existing_idx = estoque_df[
                    (estoque_df['semi'] == item_semi) & 
                    (estoque_df['gola'] == item_gola) & 
                    (estoque_df['bordado'] == item_bordado)
                ].index

                if not existing_idx.empty:
                    estoque_df.loc[existing_idx, 'quantidade'] += quantidade_adicionar
                else:
                    novo_item_estoque = pd.DataFrame([{
                        'codigo': selected_codigo, 
                        'semi': item_semi, 
                        'gola': item_gola, 
                        'bordado': item_bordado, 
                        'quantidade': quantidade_adicionar
                    }])
                    estoque_df = pd.concat([estoque_df, novo_item_estoque], ignore_index=True)
                
                salvar_estoque(estoque_df)
                st.success(f"Estoque do produto '{selected_codigo}' atualizado!")
            else:
                st.warning("Por favor, selecione um código.")

    st.subheader("Estoque Atual - Resumo")
    estoque_df = carregar_estoque()
    if not estoque_df.empty:
        # EXIBIÇÃO SIMPLIFICADA - apenas semi, gola, bordado e quantidade
        estoque_resumo = estoque_df[['semi', 'gola', 'bordado', 'quantidade']].copy()
        estoque_resumo = estoque_resumo[estoque_resumo['quantidade'] != 0]  # Ocultar itens com quantidade zero
        st.dataframe(estoque_resumo.sort_values(by=['semi', 'gola', 'bordado']))
        
        # Botão para baixar estoque
        excel_estoque = gerar_excel_formatado(estoque_resumo, "estoque_atual")
        st.download_button(
            label="📥 Baixar Estoque Atual",
            data=excel_estoque,
            file_name="estoque_atual.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Estoque vazio.")

st.markdown("---")
st.markdown("**Pure & Posh Baby** - Sistema de Relatórios v1.0")
