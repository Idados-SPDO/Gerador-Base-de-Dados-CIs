import streamlit as st
import pandas as pd
from datetime import datetime as dt
import io

def gerar_base():
    dia_atual = dt.today().strftime('%d')  # Retorna o dia atual
    numero_mes_atual = dt.today().strftime('%m')  # Retorna o número do mês atual
    ano_atual = dt.today().strftime('%Y')  # Retorna o ano atual

    arquivo_tratado = str(dia_atual) + '.' + str(numero_mes_atual) + '.' + str(ano_atual) + '.xlsx'

    for uploaded_file in uploaded_files:
        if uploaded_file.name.startswith('RelatorioCIs_'):
            st.markdown(
                """
                <style>
                .stProgress .st-bo {
                    background-color: orange;
                }
                </style>
                """, 
            unsafe_allow_html=True
            )
            progress_text = "Carregando dados..."
            progress_bar = st.progress(0, text=progress_text)
            
            progress_bar.progress(10, text=progress_text)

            df_conf_desc = pd.read_excel(uploaded_file, sheet_name='Confirmação de descrição')
            progress_bar.progress(20, text=progress_text)

            df_conf_preco = pd.read_excel(uploaded_file, sheet_name='Confirmação do preço')
            progress_bar.progress(30, text=progress_text)

            df_agenda = pd.read_excel(uploaded_file, sheet_name='Agendamento de reunião com resp')
            progress_bar.progress(40, text=progress_text)

            df_amp = pd.read_excel(uploaded_file, sheet_name='Ampliação de coleta de preços')
            progress_bar.progress(50, text=progress_text)

            df_imp = pd.read_excel(uploaded_file, sheet_name='Implantação de coleta de preço')
            progress_bar.progress(60, text=progress_text)

            df_carga = pd.read_excel(uploaded_file, sheet_name='Carga de preço')
            progress_bar.progress(70, text=progress_text)

            df_conf_desc['Tipo CI'] = 'Confirmação'
            df_conf_preco['Tipo CI'] = 'Confirmação'
            df_agenda['Tipo CI'] = 'Agendamento'
            df_amp['Tipo CI'] = 'Ampliação'
            df_imp['Tipo CI'] = 'Implantação'
            df_carga['Tipo CI'] = 'Carga'

            df_agenda = df_agenda.rename(columns={'Tipo de Insumo': 'Tipo insumo'})
            df_imp = df_imp.rename(columns={'Insumo Informado': 'Insumo informado', 'Cód. informante': 'Cód. Informante'})
            df_amp = df_amp.rename(columns={'Insumo Informado': 'Insumo informado', 'Cód. informante': 'Cód. Informante'})
            df_carga = df_carga.rename(columns={'Insumo Informado': 'Insumo informado', 'Cód. informante': 'Cód. Informante'})

            df_base = pd.concat([df_conf_desc, df_conf_preco, df_agenda, df_amp, df_imp, df_carga], ignore_index=True)
            progress_bar.progress(90, text=progress_text)

            # Salva o dataframe em um buffer de memória
            buffer = io.BytesIO()
            df_base.to_excel(buffer, index=False)
            buffer.seek(0)

            progress_bar.progress(100)

            st.success(f'A nova base foi gerada com sucesso!')

            # Adiciona o botão de download
            st.download_button(
                label='Baixar nova base',
                data=buffer,
                file_name=arquivo_tratado,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
        else:
            st.error('Nome de arquivo inválido. O arquivo selecionado deve ser iniciado em: "RelatorioCIs_"')


# Interface Streamlit
st.set_page_config(
    page_title="Nova Base CIs",
    layout="centered"
)

st.title('Gerador de Bases de Dados para o :orange[Painel de Controle de CIs]')

st.write('')
st.write('')
st.write('')
st.write('')

container = st.container()
container.subheader('Escolha o arquivo de relatório desejado:', divider="orange")
uploaded_files = container.file_uploader('', accept_multiple_files=True, type=["xlsx"])

container.write('')
container.write('')
container.write('')
container.write('')

container.caption('Pressione o botão abaixo parar gerar a nova base')

botao = container.button('Gerar nova base')

if botao:
    gerar_base()
