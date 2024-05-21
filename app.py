import streamlit as st
import pandas as pd
from datetime import datetime as dt
import os
import time

def gerar_base():
    dia_atual = dt.today().strftime('%d')  # Retorna o dia atual
    numero_mes_atual = dt.today().strftime('%m')  # Retorna o número do mês atual
    ano_atual = dt.today().strftime('%Y')  # Retorna o ano atual

    arquivo_tratado = str(dia_atual) + '.' + str(numero_mes_atual) + '.' + str(ano_atual) + '.xlsx'

    caminha_salvar = f'W:\\Planejamento e Controle - Coleta de Dados\\PC Construção e Indústria\\1. Painel de Controle SCI\\2. Dados\\{arquivo_tratado}'

    for uploaded_file in uploaded_files:
        if uploaded_file.name.startswith('RelatorioCIs_'):
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            progress_bar.progress(10)
            status_text.text('Carregando planilhas...')

            df_conf_desc = pd.read_excel(uploaded_file, sheet_name='Confirmação de descrição')
            progress_bar.progress(20)

            df_conf_preco = pd.read_excel(uploaded_file, sheet_name='Confirmação do preço')
            progress_bar.progress(30)

            df_agenda = pd.read_excel(uploaded_file, sheet_name='Agendamento de reunião com resp')
            progress_bar.progress(40)

            df_amp = pd.read_excel(uploaded_file, sheet_name='Ampliação de coleta de preços')
            progress_bar.progress(50)

            df_imp = pd.read_excel(uploaded_file, sheet_name='Implantação de coleta de preço')
            progress_bar.progress(60)

            df_carga = pd.read_excel(uploaded_file, sheet_name='Carga de preço')
            progress_bar.progress(70)

            status_text.text('Processando dados...')

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
            progress_bar.progress(90)

            df_base.to_excel(caminha_salvar, index=False)
            progress_bar.progress(100)

            st.success(f'A nova base foi gerada e salva com sucesso!')
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
    with st.spinner('Aguarde enquanto a base é gerada...'):
        gerar_base()