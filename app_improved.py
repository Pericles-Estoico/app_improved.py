import streamlit as st
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Border, Side
import os
import pickle
import json
import gspread
from google.oauth2.service_account import Credentials

# Configurações
st.set_page_config(page_title="Pure & Posh Baby - Sistema de Relatórios", page_icon="👑", layout="wide")

# Configuração Google Sheets
GOOGLE_SHEETS_CREDENTIALS = {
    "type": "service_account",
    "project_id": "relatoriodevendas",
    "private_key_id": "10b28728bffda90094580f4f739031cc153cf196",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQC2LBadtjsxtglr\n1i4IH/jqWy8uaFQYiiuxpiH/IzcIBX3XfM93eWRDUVmXJzoK2YAWas2qfkrTb1UG\n2c/cEc6xs5/8JKU0tR1VeJ//E8jJBucYCfYAxu8YCIIg5YUgDQ7agt2s+x3CYCCK\ncU5tR8FR5lnBJt+LIGY93nZu3hdVHTD0QD7nA5JCoaZYYt9FzBGnNnejbcBRYV0B\n42RsPJ+za/22uyjqLd+zBobIAlvIT+hqRde/FeyBqEreWnJ4aXc15nEYFqy1nB8E\ne8ma8vEgscbrOSlWbvTgI7aUw33AuFz8Qo0wG7mAxMVAmsXcTXXy2qAhpPaYyCcK\nGo/j6uXLAgMBAAECggEANgqPLPr9xWn0koZngmaFr3QcY350kBERFDKt/COEtD74\nzV+LpiwfN68ezi3HVKegDUZiu5Saeu2Ygh9EP8sSj3mzWJfAYInn6U6O3BsQ4b3H\n+UQfM6zQCcegXsTnwJHPGbhfrWyTL/HXRWqGcvmp2jNk5d0zzHBwlCL17D67Gja8\nmdVT218S0FZTAq8f5d+fFRcTGqBwYAamtpxc9GMPLACCODraUEuQ++hbPfDlNhpu\nkJ/2EZjea+FXE0z+U5qnrTecn6U3Lqq7XvmnuNPjXeYlVdtlwx2wNOAllDODQRSJ\nsrlGxEDTD9pjZw0/PWjXj3pLQXdQWQ7/WCRMBkAavQKBgQD60sgTGUFLtsDiKE3b\nBdMseNz4IWVJYVKOZq2SYmISS6S4H04eRqcj4Ent+JCDnZmp5cMbijSKl0xlAA9M\nobeMYB91ZlzjyKvdT+9hNcs67PoGBdP+20bn5eX13Dwi//l2ON5zU0WPL6NviANu\n2a54XRT4zGmXYYCib8TiolUe3QKBgQC57pcgOgmx3UkYmxuB1/DFdr5cM+HUnP7e\n+GWoay0S2Jww1gfmWmu3aY/OE42IueA722QQpnqiddWgnGUUdmtminDmwRMmWGvJ\n/QKQzOv6/jDd81mR+PYdT6stVH/AMGrNW/qP3VsCmcv/hqnxJ+yRii+5GW4LMcis\nusjdnt0IxwKBgEPImNdIePPsNJ4pxDiPj20yUI0iAUxeZ8AiEYBA5D4LgT1dAHCA\nKYUxhOkxxmQ7QB7BAAQ+Skq17qhQ5tGP1pmyFG5Wtn28am3Jv2hm8EBBcKQWCR+T\nxMrAv2+9D+dpg9ImNj+2XlL+zc1DVaIsY9EVXqiKHXMSn3/Gcs/IjPZlAoGACkI/\n1GdfYZD0F4d3XRKtFjgXCL9UFocTCPproX9IXWHWPFuS1ALpLpWEebpadNDMroDM\nZJ7K5Wva/aGjch2Wj3HUCOdeRx9Z0ytCmPq1ioO77oMezg8OhU+AAmBHLDN/sRUC\nHi34d4xE1TR46/Vn+B/Hwk7E45k7mUw1CQVa7MECgYEAxuEH5rmnNZiulz9N7kVO\nd9hfZxl4eFy8hsOq3WRL6dOrtSDqvGKLNvUzwsxVM52s9N1rbCbeRuBOBT68t9pI\nUPNTdbqdzkLzQ83NbpVNbRgss7dmIFiY0LVBOlUxMjTTP+v0AqePiQpeNnCRHhg4\nVtZtx0RcVARW5afen1aNLEs=\n-----END PRIVATE KEY-----\n",
    "client_email": "pureposh-service@relatoriodevendas.iam.gserviceaccount.com",
    "client_id": "101284787637223879115",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/pureposh-service%40relatoriodevendas.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

SPREADSHEET_ID = "1x_45l8hzTmWZTnjeZDB3hwr7yjHMxRrt1ii9L3c5bqU"

# Funções Google Sheets
@st.cache_resource
def init_google_sheets():
    """Inicializa conexão com Google Sheets"""
    try:
        credentials = Credentials.from_service_account_info(
            GOOGLE_SHEETS_CREDENTIALS,
            scopes=['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        )
        gc = gspread.authorize(credentials)
        return gc
    except Exception as e:
        st.error(f"Erro ao conectar com Google Sheets: {e}")
        return None

def carregar_planilha_mae_google():
    """Carrega planilha mãe do Google Sheets"""
    try:
        gc = init_google_sheets()
        if gc is None:
            return None
            
        sheet = gc.open_by_key(SPREADSHEET_ID).sheet1
        data = sheet.get_all_records()
        
        if data:
            df = pd.DataFrame(data)
            # Converter colunas numéricas
            if 'quantidade' in df.columns:
                df['quantidade'] = pd.to_numeric(df['quantidade'], errors='coerce').fillna(0)
            return df
        else:
            return None
    except Exception as e:
        st.error(f"Erro ao carregar planilha mãe: {e}")
        return None

def salvar_planilha_mae_google(df):
    """Salva planilha mãe no Google Sheets"""
    try:
        gc = init_google_sheets()
        if gc is None:
            return False
            
        sheet = gc.open_by_key(SPREADSHEET_ID).sheet1
        
        # Limpar planilha
        sheet.clear()
        
        # Adicionar cabeçalhos
        headers = df.columns.tolist()
        sheet.append_row(headers)
        
        # Adicionar dados
        for _, row in df.iterrows():
            sheet.append_row(row.tolist())
            
        return True
    except Exception as e:
        st.error(f"Erro ao salvar planilha mãe: {e}")
        return False

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

# Inicializar session_state para planilha mãe
if 'planilha_mae_carregada' not in st.session_state:
    st.session_state['planilha_mae_carregada'] = False

if 'df_mae' not in st.session_state:
    st.session_state['df_mae'] = None

# Carregar planilha mãe automaticamente do Google Sheets
if not st.session_state['planilha_mae_carregada']:
    with st.spinner("Carregando planilha mãe do Google Sheets..."):
        df_mae_google = carregar_planilha_mae_google()
        if df_mae_google is not None and not df_mae_google.empty:
            st.session_state['df_mae'] = df_mae_google
            st.session_state['planilha_mae_carregada'] = True
            st.success("✅ Planilha mãe carregada do Google Sheets!")
        else:
            st.info("ℹ️ Nenhuma planilha mãe encontrada no Google Sheets. Carregue uma planilha para começar.")

# Função para carregar Excel
def load_excel(arquivo):
    return pd.read_excel(arquivo)

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
                    'item': row['semi'],
                    'quantidade': total_semi,
                    'check': '',
                    'tipo': 'semi',
                    'categoria': row['categoria']
                })
            
            # Adicionar linha do componente
            item_componente = f"  ↳ {row['gola']} - {row['bordado']}"
            relatorio_hierarquico.append({
                'item': item_componente,
                'quantidade': row['quantidade'],
                'check': '',
                'tipo': 'componente',
                'categoria': row['categoria']
            })
        
        # Escrever dados na planilha
        row_num = 2
        for item in relatorio_hierarquico:
            ws.cell(row=row_num, column=1, value=item['item']).border = border
            ws.cell(row=row_num, column=2, value=item['quantidade']).border = border
            ws.cell(row=row_num, column=3, value=item['check']).border = border
            
            # Aplicar formatação baseada no tipo e categoria
            if item['tipo'] == 'semi':
                # Formatação para linha do semi (negrito)
                ws.cell(row=row_num, column=1).font = semi_font
                ws.cell(row=row_num, column=2).font = semi_font
                
                # Cor de fundo baseada na categoria
                if item['categoria'] == 1:  # Manga Longa
                    fill = manga_longa_fill
                elif item['categoria'] == 2:  # Manga Curta Menina
                    fill = manga_curta_menina_fill
                elif item['categoria'] == 3:  # Manga Curta Menino
                    fill = manga_curta_menino_fill
                elif item['categoria'] == 4:  # Mijão
                    fill = mijao_fill
                else:
                    fill = None
                
                if fill:
                    ws.cell(row=row_num, column=1).fill = fill
                    ws.cell(row=row_num, column=2).fill = fill
                    ws.cell(row=row_num, column=3).fill = fill
            
            row_num += 1
    
    else:
        # Relatório normal (não agrupado)
        # Adicionar colunas de ordenação
        df_ordenado = df.copy()
        df_ordenado[['categoria', 'cor_ordem', 'tamanho_ordem']] = df_ordenado['semi'].apply(
            lambda x: pd.Series(get_categoria_ordem(x))
        )
        
        # Ordenar conforme especificado
        df_ordenado = df_ordenado.sort_values([
            'categoria',      # 1=Manga Longa, 2=MC Menina, 3=MC Menino, 4=Mijão
            'cor_ordem',      # 1=Branco primeiro
            'tamanho_ordem',  # 1=RN, 2=P, 3=M, 4=G
            'semi',
            'gola',
            'bordado'
        ])
        
        # Remover colunas de ordenação para o relatório final
        colunas_originais = [col for col in df_ordenado.columns if col not in ['categoria', 'cor_ordem', 'tamanho_ordem']]
        df_ordenado = df_ordenado[colunas_originais]
        
        # Cabeçalhos
        for col_num, column_title in enumerate(df_ordenado.columns, 1):
            cell = ws.cell(row=1, column=col_num, value=column_title)
            cell.fill = header_fill
            cell.font = header_font
            cell.border = border
        
        # Dados
        for row_num, row_data in enumerate(df_ordenado.itertuples(index=False), 2):
            for col_num, value in enumerate(row_data, 1):
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
st.markdown("---")

# Seção 1: Upload da Planilha Mãe
st.header("📊 1. Planilha Mãe")

if st.session_state['planilha_mae_carregada'] and st.session_state['df_mae'] is not None:
    st.success(f"✅ Planilha mãe carregada! ({len(st.session_state['df_mae'])} registros)")
    
    # Mostrar preview da planilha mãe
    with st.expander("👀 Visualizar Planilha Mãe"):
        st.dataframe(st.session_state['df_mae'])
else:
    st.info("📤 Carregue a planilha mãe para começar")

# Upload da planilha mãe
planilha_mae = st.file_uploader("Selecione a planilha mãe", type=['xlsx', 'xls'], key="planilha_mae")

if planilha_mae is not None:
    try:
        df_mae = load_excel(planilha_mae)
        st.session_state['df_mae'] = df_mae
        st.session_state['planilha_mae_carregada'] = True
        
        # Salvar no Google Sheets
        with st.spinner("Salvando planilha mãe no Google Sheets..."):
            if salvar_planilha_mae_google(df_mae):
                st.success("✅ Planilha mãe salva no Google Sheets com sucesso!")
            else:
                st.warning("⚠️ Erro ao salvar no Google Sheets, mas planilha carregada na sessão.")
        
        st.success(f"✅ Planilha mãe carregada! ({len(df_mae)} registros)")
        
        # Mostrar preview
        with st.expander("👀 Preview da Planilha Mãe"):
            st.dataframe(df_mae.head())
            
    except Exception as e:
        st.error(f"❌ Erro ao carregar planilha mãe: {e}")

st.markdown("---")

# Seção 2: Upload do Relatório de Faltantes (para atualizar planilha mãe)
st.header("📋 2. Relatório de Faltantes (Atualização)")

st.info("💡 **Importante:** Quando você subir um relatório de faltantes preenchido, os novos códigos serão adicionados permanentemente à planilha mãe.")

relatorio_faltantes = st.file_uploader("Selecione o relatório de faltantes preenchido", type=['xlsx', 'xls'], key="faltantes")

if relatorio_faltantes is not None and st.session_state['planilha_mae_carregada']:
    try:
        df_faltantes = load_excel(relatorio_faltantes)
        st.success(f"✅ Relatório de faltantes carregado! ({len(df_faltantes)} registros)")
        
        # Mostrar preview
        with st.expander("👀 Preview do Relatório de Faltantes"):
            st.dataframe(df_faltantes.head())
        
        # Botão para atualizar planilha mãe
        if st.button("🔄 Atualizar Planilha Mãe com Faltantes", type="primary"):
            try:
                # Combinar planilha mãe com faltantes
                df_mae_atualizada = pd.concat([st.session_state['df_mae'], df_faltantes], ignore_index=True)
                
                # Remover duplicatas se existirem
                if 'semi' in df_mae_atualizada.columns and 'gola' in df_mae_atualizada.columns and 'bordado' in df_mae_atualizada.columns:
                    df_mae_atualizada = df_mae_atualizada.drop_duplicates(subset=['semi', 'gola', 'bordado'], keep='last')
                
                # Atualizar session state
                st.session_state['df_mae'] = df_mae_atualizada
                
                # Salvar no Google Sheets
                with st.spinner("Salvando planilha mãe atualizada no Google Sheets..."):
                    if salvar_planilha_mae_google(df_mae_atualizada):
                        st.success("✅ Planilha mãe atualizada e salva no Google Sheets com sucesso!")
                        st.success(f"📊 Total de registros: {len(df_mae_atualizada)}")
                        st.balloons()
                    else:
                        st.error("❌ Erro ao salvar no Google Sheets")
                        
            except Exception as e:
                st.error(f"❌ Erro ao atualizar planilha mãe: {e}")
                
elif relatorio_faltantes is not None and not st.session_state['planilha_mae_carregada']:
    st.warning("⚠️ Carregue primeiro a planilha mãe para poder atualizá-la com os faltantes.")

st.markdown("---")

# Seção 3: Upload das Planilhas de Pedidos
st.header("📦 3. Planilhas de Pedidos")

if not st.session_state['planilha_mae_carregada']:
    st.warning("⚠️ Carregue primeiro a planilha mãe para processar os pedidos.")
else:
    planilhas_pedidos = st.file_uploader("Selecione as planilhas de pedidos", type=['xlsx', 'xls'], accept_multiple_files=True)
    
    if planilhas_pedidos:
        st.success(f"✅ {len(planilhas_pedidos)} planilha(s) de pedidos carregada(s)")
        
        # Processar todas as planilhas
        todos_pedidos = []
        for planilha in planilhas_pedidos:
            try:
                df_pedido = load_excel(planilha)
                df_pedido['arquivo_origem'] = planilha.name
                todos_pedidos.append(df_pedido)
            except Exception as e:
                st.error(f"❌ Erro ao carregar {planilha.name}: {e}")
        
        if todos_pedidos:
            # Combinar todos os pedidos
            df_todos_pedidos = pd.concat(todos_pedidos, ignore_index=True)
            st.success(f"✅ Total de {len(df_todos_pedidos)} itens processados")
            
            # Mostrar preview
            with st.expander("👀 Preview dos Pedidos"):
                st.dataframe(df_todos_pedidos.head())
            
            st.markdown("---")
            
            # Seção 4: Gerar Relatórios
            st.header("📊 4. Gerar Relatórios")
            
            if st.button("🚀 Gerar Todos os Relatórios", type="primary"):
                try:
                    # Preparar dados para relatórios
                    df_mae = st.session_state['df_mae']
                    
                    # Verificar se as colunas necessárias existem
                    colunas_necessarias = ['semi', 'gola', 'bordado']
                    if not all(col in df_mae.columns for col in colunas_necessarias):
                        st.error(f"❌ Planilha mãe deve conter as colunas: {colunas_necessarias}")
                        st.stop()
                    
                    if not all(col in df_todos_pedidos.columns for col in colunas_necessarias):
                        st.error(f"❌ Planilhas de pedidos devem conter as colunas: {colunas_necessarias}")
                        st.stop()
                    
                    # Criar chave única para merge
                    df_mae['chave'] = df_mae['semi'].astype(str) + '|' + df_mae['gola'].astype(str) + '|' + df_mae['bordado'].astype(str)
                    df_todos_pedidos['chave'] = df_todos_pedidos['semi'].astype(str) + '|' + df_todos_pedidos['gola'].astype(str) + '|' + df_todos_pedidos['bordado'].astype(str)
                    
                    # 1. Relatório de Itens Existentes
                    st.subheader("📋 1. Relatório de Itens Existentes")
                    itens_existentes = df_todos_pedidos[df_todos_pedidos['chave'].isin(df_mae['chave'])]
                    
                    if not itens_existentes.empty:
                        excel_existentes = gerar_excel_formatado(itens_existentes.drop('chave', axis=1), "Itens_Existentes")
                        st.download_button(
                            label="📥 Download - Itens Existentes",
                            data=excel_existentes,
                            file_name="01_Itens_Existentes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"✅ {len(itens_existentes)} itens existentes encontrados")
                    else:
                        st.info("ℹ️ Nenhum item existente encontrado")
                    
                    # 2. Relatório de Itens Faltantes
                    st.subheader("📋 2. Relatório de Itens Faltantes")
                    itens_faltantes = df_todos_pedidos[~df_todos_pedidos['chave'].isin(df_mae['chave'])]
                    
                    if not itens_faltantes.empty:
                        excel_faltantes = gerar_excel_formatado(itens_faltantes.drop('chave', axis=1), "Itens_Faltantes")
                        st.download_button(
                            label="📥 Download - Itens Faltantes",
                            data=excel_faltantes,
                            file_name="02_Itens_Faltantes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success(f"✅ {len(itens_faltantes)} itens faltantes encontrados")
                    else:
                        st.info("ℹ️ Nenhum item faltante encontrado")
                    
                    # 3. Relatório de Componentes Existentes
                    st.subheader("📋 3. Relatório de Componentes Existentes")
                    if not itens_existentes.empty:
                        excel_comp_existentes = gerar_excel_formatado(itens_existentes.drop('chave', axis=1), "Componentes_Existentes", agrupar_por_semi=True)
                        st.download_button(
                            label="📥 Download - Componentes Existentes",
                            data=excel_comp_existentes,
                            file_name="03_Componentes_Existentes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("✅ Relatório de componentes existentes gerado")
                    else:
                        st.info("ℹ️ Nenhum componente existente para gerar relatório")
                    
                    # 4. Relatório de Componentes Faltantes
                    st.subheader("📋 4. Relatório de Componentes Faltantes")
                    if not itens_faltantes.empty:
                        excel_comp_faltantes = gerar_excel_formatado(itens_faltantes.drop('chave', axis=1), "Componentes_Faltantes", agrupar_por_semi=True)
                        st.download_button(
                            label="📥 Download - Componentes Faltantes",
                            data=excel_comp_faltantes,
                            file_name="04_Componentes_Faltantes.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("✅ Relatório de componentes faltantes gerado")
                    else:
                        st.info("ℹ️ Nenhum componente faltante para gerar relatório")
                    
                    st.markdown("---")
                    st.success("🎉 **Todos os relatórios foram gerados com sucesso!**")
                    st.balloons()
                    
                except Exception as e:
                    st.error(f"❌ Erro ao gerar relatórios: {e}")
                    st.exception(e)

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666; font-size: 0.8em;'>
    <p>👑 Pure & Posh Baby - Sistema de Relatórios v2.0</p>
    <p>💾 Planilha mãe salva permanentemente no Google Sheets</p>
</div>
""", unsafe_allow_html=True)

