# -*- coding: utf-8 -*-
"""
Sistema Web Avançado para Formatar e Criar Nomes de Pastas.

Versão 4.0:
- Adicionada sugestão de modelos com separadores (_ e -).
- Implementada ordenação automática dos dados por data crescente antes da geração.
- Adicionada verificação e criação do diretório base (pai) caso ele não exista.
- Interface do usuário totalmente traduzida para o português brasileiro.
- Corrigido o erro 'KeyError' que ocorria quando um campo não era mapeado pelo usuário.
- Adicionada limpeza automática de separadores duplicados ou finais.

Como executar:
1. Salve este arquivo como `app.py`.
2. Instale as bibliotecas necessárias:
   pip install streamlit pandas openpyxl
3. No terminal, execute o comando:
   streamlit run app.py
"""
import streamlit as st
import pandas as pd
import os
import re

# --- Funções de Lógica ---

def processar_dados(df, mapeamento, template):
    """
    Processa o DataFrame para gerar os nomes das pastas com base no mapeamento e modelo do usuário.
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
                dt_inicio = pd.to_datetime(row[mapeamento['data_inicio']], dayfirst=True)
                partes_nome['DATA'] = dt_inicio.strftime('%d-%m-%Y')
                partes_nome['HORA_INICIO'] = dt_inicio.strftime('%H-%M-%S')
            
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
    page_title="Criador de Pastas a partir de Planilha",
    page_icon="📂",
    layout="wide"
)

# --- Interface do Usuário ---

st.title("⚙️ Criador de Pastas a partir de Planilha")
st.markdown("Uma ferramenta flexível para gerar nomes de pastas e criá-las diretamente no seu computador.")

# --- Passo 1: Upload ---
st.header("Passo 1: Envie sua Planilha")
uploaded_file = st.file_uploader(
    "Arraste e solte o arquivo Excel (.xlsx) aqui ou clique para procurar",
    type=["xlsx"],
    label_visibility="collapsed"
)

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, engine='openpyxl')
        st.success("Planilha carregada com sucesso!")
        st.subheader("Pré-visualização dos dados:")
        st.dataframe(df.head(), use_container_width=True)
        
        colunas_disponiveis = ["N/A"] + df.columns.tolist()

        # --- Passo 2: Mapeamento e Modelo ---
        st.header("Passo 2: Configure a Conversão")
        
        col1, col2 = st.columns(2)

        with col1:
            with st.expander("Mapeamento de Colunas", expanded=True):
                st.info("Associe os campos do sistema às colunas da sua planilha. Deixe como 'N/A' se não quiser usar um campo.")
                mapeamento = {
                    'data_inicio': st.selectbox("Coluna para Data e Hora de Início (Obrigatório para Ordenação)", colunas_disponiveis, key='map_di'),
                    'data_fim': st.selectbox("Coluna para Data e Hora de Fim", colunas_disponiveis, key='map_df'),
                    'condutor': st.selectbox("Coluna para Nome do Condutor", colunas_disponiveis, key='map_c'),
                    'cpf': st.selectbox("Coluna para CPF", colunas_disponiveis, key='map_cpf'),
                    'maquina': st.selectbox("Coluna para Máquina/Equipamento", colunas_disponiveis, key='map_m'),
                }
        
        with col2:
            with st.expander("Modelo do Nome da Pasta", expanded=True):
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
                
                template_usuario = st.text_input(
                    "Edite seu modelo aqui:",
                    value=sugestoes[sugestao_selecionada]
                )

        # --- Passo 3: Geração ---
        st.header("Passo 3: Gerar e Criar Pastas")

        if st.button("▶️ Gerar Nomes das Pastas"):
            if mapeamento['data_inicio'] != 'N/A':
                try:
                    # Ordena o dataframe pela data de início
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
                st.text_area(
                    "Nomes gerados (em ordem cronológica):",
                    "\n".join(nomes_gerados),
                    height=250
                )
                st.download_button(
                    label="📥 Baixar Lista de Nomes (.txt)",
                    data="\n".join(nomes_gerados),
                    file_name="nomes_de_pastas.txt",
                    mime="text/plain"
                )

        # --- Passo 4: Criação das Pastas ---
        if 'nomes_gerados' in st.session_state and st.session_state['nomes_gerados']:
            st.markdown("---")
            st.subheader("Opcional: Criar Pastas no seu Computador")
            
            st.warning("**Atenção:** Esta função criará pastas reais no diretório que você especificar.")
            
            caminho_diretorio = st.text_input("Cole aqui o caminho completo do diretório onde as pastas devem ser criadas (ex: C:\\Usuários\\SeuUsuario\\Documentos\\Relatorios)")

            if st.button("🚀 Criar Pastas no Diretório Acima"):
                if caminho_diretorio:
                    try:
                        # Verifica se o diretório base existe. Se não, informa que será criado.
                        if not os.path.isdir(caminho_diretorio):
                            st.info(f"O diretório base '{caminho_diretorio}' não existe e será criado.")
                        
                        pastas_criadas = 0
                        erros_criacao = []
                        with st.spinner(f"Criando pastas em '{caminho_diretorio}'..."):
                            for nome_pasta in st.session_state['nomes_gerados']:
                                try:
                                    nome_pasta_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nome_pasta)
                                    caminho_completo = os.path.join(caminho_diretorio, nome_pasta_sanitizado)
                                    os.makedirs(caminho_completo, exist_ok=True)
                                    pastas_criadas += 1
                                except Exception as e:
                                    erros_criacao.append(f"Falha ao criar '{nome_pasta}': {e}")
                        
                        st.success(f"Operação concluída! {pastas_criadas} pastas foram criadas/verificadas com sucesso em '{caminho_diretorio}'.")
                        if erros_criacao:
                            st.error("Alguns erros ocorreram durante a criação:")
                            st.json(erros_criacao)
                    except Exception as e:
                        st.error(f"Erro ao processar o caminho do diretório: {e}")
                else:
                    st.error("O caminho do diretório não pode estar vazio. Por favor, especifique um local.")
    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}. Verifique se o arquivo não está corrompido.")

