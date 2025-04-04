import streamlit as st
import base64
import time
import pandas as pd
import os
from pathlib import Path
from processar_extrato import process_pdf_to_excel

# Configuração da página - DEVE SER A PRIMEIRA CHAMADA STREAMLIT
st.set_page_config(
    page_title="Conversor de Extrato - Financeiro",
    page_icon="💼",
    layout="centered",
    initial_sidebar_state="collapsed"
)

# CSS mínimo para o botão de download
st.markdown("""
<style>
    .custom-download-btn {
        background-color: #00326D;
        color: white;
        padding: 12px 24px;
        text-align: center;
        text-decoration: none;
        font-size: 16px;
        border-radius: 4px;
        border: none;
        cursor: pointer;
        display: inline-block;
        font-weight: 500;
    }
    
    .custom-download-btn:hover {
        background-color: #0072CE;
        color: white;
    }
    
    .download-container {
        display: flex;
        justify-content: center;
        margin: 20px 0;
    }
</style>
""", unsafe_allow_html=True)

def main():
    # Caminho da logo que funciona tanto localmente quanto no Streamlit Cloud
    current_dir = Path(__file__).parent if "__file__" in locals() else Path.cwd()
    logo_path = current_dir / "logo.png"
    
    # Configuração do cabeçalho com logo
    if os.path.exists(logo_path):
        # Layout em duas colunas para o cabeçalho
        col1, col2 = st.columns([1, 3])
        
        # Coluna da logo
        with col1:
            # Espaço adicional para mover a logo para baixo
            st.write("")
            st.write("")
            st.image(str(logo_path), width=500)
        
        # Coluna do título
        with col2:
            st.title("Conversor de Extratos")
            st.write("Departamento Financeiro")
    else:
        # Cabeçalho sem logo caso a imagem não seja encontrada
        st.title("Conversor de Extratos")
        st.write("Departamento Financeiro")
        st.warning("Logo não encontrada. A visualização continuará sem a imagem.")
    
    st.markdown("---")
    
    # Seção de instruções
    st.subheader("📋 Como utilizar")
    st.write("Esta ferramenta converte extratos de pagamento em formato PDF para planilhas Excel organizadas.")
    st.markdown("""
    1. Faça upload do arquivo PDF contendo o extrato bancário
    2. Clique em "Processar Extrato"
    3. Após o processamento, baixe o arquivo Excel gerado
    """)
    
    st.markdown("---")
    
    # Seção de upload
    st.subheader("Selecionar Arquivo")
    uploaded_file = st.file_uploader("Escolha o arquivo PDF de extrato bancário", type="pdf")
    
    if uploaded_file:
        st.success(f"Arquivo carregado: {uploaded_file.name}")
        
        # Botão para iniciar processamento
        if st.button("Processar Extrato", type="primary"):
            # Mostrar barra de progresso
            with st.spinner("Processando o extrato... Por favor, aguarde..."):
                # Simular um pequeno atraso para mostrar o spinner
                progress_bar = st.progress(0)
                
                # Atualiza barra de progresso
                for i in range(100):
                    time.sleep(0.02)
                    progress_bar.progress(i + 1)
                
                # Processar o arquivo
                excel_data, num_transactions, error = process_pdf_to_excel(uploaded_file)
                
                # Verificar se há erro
                if error:
                    st.error(error)
                else:
                    # Criar link de download para o Excel
                    b64 = base64.b64encode(excel_data).decode()
                    filename = uploaded_file.name.replace('.pdf', '.xlsx')
                    
                    # Mensagem de sucesso
                    st.success(f"✅ Processamento concluído! Foram encontradas {num_transactions} transações no extrato.")
                    
                    # Botão de download
                    st.markdown(f"""
                    <div class="download-container">
                        <a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" 
                           download="{filename}" 
                           class="custom-download-btn">
                           📥 Baixar Planilha Excel
                        </a>
                    </div>
                    """, unsafe_allow_html=True)
                    
                    # Exibir prévia dos dados
                    st.subheader("Prévia das transações")
                    df = pd.read_excel(excel_data)
                    st.dataframe(df.head(5), use_container_width=True)
    
    # Rodapé
    st.markdown("---")
    st.caption("© 2025 Departamento Financeiro - Ferramenta Interna | Versão 1.0")

if __name__ == "__main__":
    main()