# -*- coding: utf-8 -*-
"""
Sistema Web Avançado para Formatar e Criar Nomes de Pastas.

Versão 5.9:
- Adicionada verificação explícita de permissão de escrita no diretório de destino.
- Melhorada a lógica de criação de pastas para ser mais robusta e dar feedback detalhado.
- Adicionadas mensagens de erro específicas para problemas de permissão em servidores.
- Substituída a validação de caminho padrão por uma função mais robusta,
  específica para Windows Server e caminhos de rede (UNC).
- Implementada a criação de subpastas por mês (ex: 06-Junho, 07-Julho).
- Implementado o mapeamento automático e inteligente de colunas.

Como executar:
1. Salve este ficheiro como `app.py`.
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

def is_windows_abs_path(path):
    """
    Valida de forma mais robusta se um caminho é absoluto no Windows,
    verificando por letras de unidade (C:\) ou caminhos de rede UNC (\\servidor).
    Esta função é mais fiável em ambientes de servidor.
    """
    path = path.strip('"') # Remove aspas que podem vir do 'copiar como caminho'
    if re.match(r'^[a-zA-Z]:[\\/]', path):
        return True
    if path.startswith('\\\\'):
        return True
    return False

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
    Processa o DataFrame para gerar os nomes das pastas e retorna uma lista de tuplos
    contendo (nome_final, objeto_datetime_inicio).
    """
    items_gerados = []
    erros = []

    for index, row in df.iterrows():
        try:
            partes_nome = {
                'DATA': '', 'HORA_INICIO': '', 'HORA_FIM': '',
                'CONDUTOR': '', 'CPF': '', 'MAQUINA': ''
            }
            dt_inicio_obj = None

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
            
            items_gerados.append((nome_final, dt_inicio_obj))

        except Exception as e:
            erros.append(f"Erro na linha {index + 2} da planilha: {e}")

    return items_gerados, erros

# --- Configuração da Página ---
st.set_page_config(
    page_title="Criador de Pastas a partir de Planilha",
    page_icon="📂",
    layout="wide"
)

# --- Interface do Usuário ---
st.title("⚙️ Criador de Pastas a partir de Planilha")
st.markdown("Uma ferramenta flexível para gerar nomes de pastas e criá-las diretamente no seu computador.")

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
                template_usuario = st.text_input("Edite seu modelo aqui:", value=sugestoes[sugestao_selecionada])

        st.header("Passo 3: Gerar e Criar Pastas")

        if st.button("▶️ Gerar Nomes das Pastas"):
            if mapeamento['data_inicio'] != 'N/A':
                try:
                    df_ordenado = df.sort_values(by=mapeamento['data_inicio']).copy()
                    st.info("Os dados foram ordenados pela data de início em ordem crescente.")
                    items_gerados, erros = processar_dados(df_ordenado, mapeamento, template_usuario)
                except Exception as e:
                    st.error(f"Erro ao tentar ordenar pela coluna de data: {e}")
                    items_gerados, erros = [], []
            else:
                st.warning("A coluna de Data de Início não foi selecionada. Os dados não serão ordenados.")
                items_gerados, erros = processar_dados(df, mapeamento, template_usuario)
            
            if erros:
                st.warning("Ocorreram alguns erros durante o processamento:")
                st.json(erros)
            
            if items_gerados:
                st.session_state['items_gerados'] = items_gerados
                nomes_para_exibir = [item[0] for item in items_gerados]
                st.text_area("Nomes gerados (em ordem cronológica):", "\n".join(nomes_para_exibir), height=250)
                st.download_button("📥 Baixar Lista de Nomes (.txt)", "\n".join(nomes_para_exibir), "nomes_de_pastas.txt", "text/plain")

        if 'items_gerados' in st.session_state and st.session_state['items_gerados']:
            st.markdown("---")
            st.subheader("Opcional: Criar Pastas no seu Computador")
            st.info("As pastas serão criadas dentro de subpastas com o nome do mês (ex: 06-Junho, 07-Julho).")
            
            with st.expander("Como selecionar o diretório de destino?", expanded=True):
                st.markdown("""
                1. No seu computador, abra o **Explorador de Ficheiros** e navegue até à pasta onde quer salvar.
                2. Clique na barra de endereço na parte de cima da janela.
                3. O caminho completo será selecionado (ex: `C:\\Utilizadores\\SeuNome\\Documentos`).
                4. Copie o caminho (**Ctrl+C**).
                5. Cole o caminho no campo abaixo (**Ctrl+V**).
                """)
            
            caminho_diretorio = st.text_input("Cole aqui o caminho completo do diretório de destino:")
            
            if caminho_diretorio:
                caminho_limpo = caminho_diretorio.strip().strip('"').strip("'")
                
                st.success(f"Diretório de destino definido: `{caminho_limpo}`")
                if st.button("🚀 Criar Pastas no Diretório Definido"):
                    try:
                        if not is_windows_abs_path(caminho_limpo):
                             st.error("O caminho fornecido não parece ser um caminho absoluto válido para Windows. Verifique se começa com uma letra de unidade (ex: C:\\) ou é um caminho de rede (ex: \\\\servidor\\pasta).")
                        else:
                            # **NOVA VALIDAÇÃO DE PERMISSÃO**
                            # Tenta criar o diretório base para verificar se existe e se temos permissão
                            st.write(f"Verificando o diretório base: `{caminho_limpo}`...")
                            os.makedirs(caminho_limpo, exist_ok=True)
                            
                            if not os.access(caminho_limpo, os.W_OK):
                                raise PermissionError("Sem permissão de escrita.")

                            st.write("Verificação de permissão bem-sucedida. A criar pastas...")
                            meses = {
                                1: "01-Janeiro", 2: "02-Fevereiro", 3: "03-Março", 4: "04-Abril",
                                5: "05-Maio", 6: "06-Junho", 7: "07-Julho", 8: "08-Agosto",
                                9: "09-Setembro", 10: "10-Outubro", 11: "11-Novembro", 12: "12-Dezembro"
                            }
                            pastas_criadas = 0
                            erros_criacao = []
                            with st.spinner(f"Criando pastas em '{caminho_limpo}'..."):
                                for nome_pasta, data_inicio_obj in st.session_state['items_gerados']:
                                    try:
                                        if data_inicio_obj is None:
                                            erros_criacao.append(f"Não foi possível criar '{nome_pasta}': Data de início não fornecida para determinar o mês.")
                                            continue
                                        
                                        mes_numero = data_inicio_obj.month
                                        nome_mes = meses.get(mes_numero, "Mes_Desconhecido")
                                        diretorio_mes = os.path.join(caminho_limpo, nome_mes)
                                        
                                        nome_pasta_sanitizado = re.sub(r'[<>:"/\\|?*]', '', nome_pasta)
                                        caminho_completo = os.path.join(diretorio_mes, nome_pasta_sanitizado)
                                        os.makedirs(caminho_completo, exist_ok=True)
                                        pastas_criadas += 1
                                    except Exception as e:
                                        erros_criacao.append(f"Falha ao criar '{nome_pasta}': {e}")
                            
                            st.success(f"Operação concluída! {pastas_criadas} pastas foram criadas/verificadas com sucesso.")
                            if erros_criacao:
                                st.error("Alguns erros ocorreram durante a criação:")
                                st.json(erros_criacao)

                    except PermissionError:
                        st.error(f"**Erro de Permissão!** O script não tem permissão para criar pastas no diretório '{caminho_limpo}'. Por favor, verifique as permissões da pasta para o utilizador que está a executar o script, ou tente executar como administrador.")
                    except FileNotFoundError:
                        st.error(f"**Caminho não encontrado!** O diretório base '{caminho_limpo}' não existe ou não é acessível. Por favor, verifique se o caminho está correto.")
                    except Exception as e:
                        st.error(f"Ocorreu um erro inesperado: {e}")

    except Exception as e:
        st.error(f"Ocorreu um erro ao ler o arquivo Excel: {e}. Verifique se o arquivo não está corrompido.")

