# -*- coding: utf-8 -*-
"""
Sistema Web Avançado para Formatar Nomes de Pastas.

Versão 6.1 (Versão Web):
- Removida toda a funcionalidade de criação de pastas locais para garantir a
  compatibilidade com o Streamlit Community Cloud.
- A aplicação foca-se na geração da lista de nomes e no download do ficheiro .txt.
- Implementado o mapeamento automático e inteligente de colunas.
- Adicionada sugestão de modelos com separadores (_ e -).
- Implementada ordenação automática dos dados por data crescente antes da geração.

Como executar:
1. Salve este ficheiro como `app.py`.
2. Crie um ficheiro `requirements.txt` com o seguinte conteúdo:
   streamlit
   pandas
   openpyxl
3. No terminal, execute o comando:
   streamlit run app.py
"""
import streamlit as st
import pandas as pd
import re

# --- Funções de Lógica ---

def guess_mappings(columns):
    """
    Tenta adivinhar o mapeamento das colunas com base em nomes e palavras-chave comuns.
    Retorna um dicionário com os nomes das colunas adivinhadas.
    """
    mapping_keywords = {
        'data_inicio': ['data início', 'datainicio', 'data_inicio', 'start date', 'inicio', 'começo'],
        'data_fim': ['data fim', 'datafim', 'data_fim', 'end date', 'fim', 'término', 'termino'],
        'condutor': ['condutor', 'motorista', 'driver', 'nome', 'operador'],
        'cpf': ['cpf'],
        'maquina': ['maquina', 'máquina', 'equipamento', 'equipment', 'veiculo', 'viatura']
    }
    
    guessed_map = {}
    normalized_columns = {col: re.sub(r'[^a-z0-9]', '', col.lower()) for col in columns}
    
    for map_key, keywords in mapping_keywords.items():
        found = False
        for col, normalized_col in normalized_columns.items():
            for keyword in keywords:
                normalized_keyword = re.sub(r'[^a-z0-9]', '', keyword.lower())
                if normalized_keyword in normalized_col:
                    guessed_map[map_key] = col
                    found = True
                    break
            if found:
                break
    
    return guessed_map

def processar_dados(df, mapeamento, template):
    """
    Processa o DataFrame para gerar os nomes das pastas e retorna uma lista de nomes.
    """
    nomes_gerados = []
    erros = []

    for index, row in df.iterrows():
        try:
            partes_nome = {
                'DATA': '', 'HORA_INICIO': '', 'HORA_FIM': '',
                'CONDUTOR': '', 'CPF': '', 'MAQUINA': ''
            }

            if mapeamento['data_inicio'] != "N/A":
                dt_inicio_obj = pd.to_datetime(row[mapeamento['data_inicio']], dayfirst=True)
                partes_nome['DATA'] = dt_inicio_obj.strftime('%d-%m-%Y')
                partes_nome['HORA_INICIO'] = dt_inicio_obj.strftime('%H-%M-%S')
            
            if mapeamento['data_fim'] != "N/A":
                dt_fim = pd.to_datetime(row[mapeamento['data_fim']], dayfirst=True)
                partes_nome['HORA_FIM'] = dt_fim.strftime('%H-%M-%S')

            if mapeamento['condutor'] != "N/A":
                partes_nome['CONDUTOR'] = str(row[mapeamento['condutor']]).strip().replace(' ', '-')

            if mapeamento['cpf'] != "N/A":
                partes_nome['CPF'] = str(row[mapeamento['cpf']]).split('.')[0]

            if mapeamento['maquina'] != "N/A":
                partes_nome['MAQUINA'] = str(row[mapeamento['maquina']]).strip()

            nome_final = template.format(**partes_nome)
            nome_final = re.sub(r'[_]+', '_', nome_final)
            nome_final = re.sub(r'[-]+', '-', nome_final)
            nome_final = nome_final.strip('_- ')
            
            nomes_gerados.append(nome_final)

        except Exception as e:
            erros.append(f"Erro na linha {index + 2} da planilha: {e}")

    return nomes_gerados, erros

# --- Configuração da Página ---
st.set_page_config(
    page_title="Gerador de Nomes a partir de Planilha",
    page_icon="📄",
    layout="wide"
)

# --- Interface do Usuário ---
st.title("⚙️ Gerador de Nomes a partir de Planilha")
st.markdown("Uma ferramenta para converter os dados da sua planilha em nomes formatados.")

st.header("Passo 1: Envie sua Planilha")
uploaded_file = st.file_uploader(
    "Arraste e solte o arquivo Excel (.xlsx) aqui ou clique para procurar",
    type=["xlsx"],
    label_visibility="collapsed"
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        df.columns = [str(col).strip() for col in df.columns]
        
        st.success("Planilha carregada com sucesso!")
        st.subheader("Pré-visualização dos dados:")
        st.dataframe(df.head(), use_container_width=True)
        
        colunas_disponiveis = ["N/A"] + df.columns.tolist()

        st.header("Passo 2: Configure a Conversão")
        
        guessed_map = guess_mappings(df.columns)
        st.info("O sistema tentou adivinhar o mapeamento das colunas abaixo. Por favor, verifique se está correto.")
        
        col1, col2 = st.columns(2)

        with col1:
            with st.expander("Mapeamento de Colunas", expanded=True):
                mapeamento = {}
                def get_col_index(key):
                    col_name = guessed_map.get(key, 'N/A')
                    return colunas_disponiveis.index(col_name) if col_name in colunas_disponiveis else 0

                mapeamento['data_inicio'] = st.selectbox("Coluna para Data e Hora de Início (Obrigatório para Ordenação)", colunas_disponiveis, index=get_col_index('data_inicio'), key='map_di')
                mapeamento['data_fim'] = st.selectbox("Coluna para Data e Hora de Fim", colunas_disponiveis, index=get_col_index('data_fim'), key='map_df')
                mapeamento['condutor'] = st.selectbox("Coluna para Nome do Condutor", colunas_disponiveis, index=get_col_index('condutor'), key='map_c')
                mapeamento['cpf'] = st.selectbox("Coluna para CPF", colunas_disponiveis, index=get_col_index('cpf'), key='map_cpf')
                mapeamento['maquina'] = st.selectbox("Coluna para Máquina/Equipamento", colunas_disponiveis, index=get_col_index('maquina'), key='map_m')
        
        with col2:
            with st.expander("Modelo do Nome", expanded=True):
                st.info("Escolha uma sugestão ou edite o modelo livremente usando as variáveis abaixo.")
                st.code("{DATA} {HORA_INICIO} {HORA_FIM} {CONDUTOR} {CPF} {MAQUINA}")
                sugestoes = {
                    "Padrão (Underline)": "{DATA}_{CONDUTOR}_{CPF}_{MAQUINA}",
                    "Completo (Underline)": "{DATA}_{HORA_INICIO}_{HORA_FIM}_{CONDUTOR}_{CPF}_{MAQUINA}",
                    "Padrão (Hífen)": "{DATA}-{CONDUTOR}-{CPF}-{MAQUINA}",
                    "Completo (Hífen)": "{DATA}-{HORA_INICIO}-{HORA_FIM}-{CONDUTOR}-{CPF}-{MAQUINA}",
                    "Apenas Data e Condutor": "{DATA}_{CONDUTOR}",
                }
                sugestao_selecionada = st.selectbox("Sugestões de Modelo:", list(sugestoes.keys()))
                template_usuario = st.text_input("Edite seu modelo aqui:", value=sugestoes[sugestao_selecionada])

        st.header("Passo 3: Gerar Lista de Nomes")

        if st.button("▶️ Gerar Nomes"):
            if mapeamento['data_inicio'] != 'N/A':
                try:
                    df_ordenado = df.sort_values(by=mapeamento['data_inicio']).copy()
                    st.info("Os dados foram ordenados pela data de início em ordem crescente.")
                    nomes_gerados, erros = processar_dados(df_ordenado, mapeamento, template_usuario)
                except Exception as e:
                    st.error(f"Erro ao tentar ordenar pela coluna de data: {e}")
                    nomes_gerados, erros = [], []
            else:
                st.warning("A coluna de Data de Início não foi selecionada. Os dados não serão ordenados.")
                nomes_gerados, erros = processar_dados(df, mapeamento, template_usuario)
            
            if erros:
                st.warning("Ocorreram alguns erros durante o processamento:")
                st.json(erros)
            
            if nomes_gerados:
                st.session_state['nomes_gerados'] = nomes_gerados
                st.text_area("Nomes gerados (em ordem cronológica):", "\n".join(nomes_gerados), height=300)
                st.download_button(
                    label="📥 Baixar Lista de Nomes (.txt)",
                    data="\n".join(nomes_gerados),
                    file_name="nomes_formatados.txt",
                    mime="text/plain"
                )

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}. Verifique se o arquivo não está corrompido.")

