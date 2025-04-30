import streamlit as st
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import io

# -------------------------
# Classes POO e funções
# -------------------------
class Produto:
    def __init__(self, nome, volume_total_ml, preco_unitario_ml):
        self.nome = nome
        self.volume_total_ml = volume_total_ml
        self.preco_unitario_ml = preco_unitario_ml
        self.nivel_critico_frac = 0.05  # 5%

    def usar(self, quantidade):
        if quantidade > self.volume_total_ml:
            raise ValueError(f"Estoque insuficiente para {self.nome}")
        self.volume_total_ml -= quantidade

    def em_nivel_critico(self):
        return self.volume_total_ml <= self.nivel_critico_frac * self.volume_total_ml_initial

    def set_initial(self):
        # guardar volume inicial para cálculo de crítico
        self.volume_total_ml_initial = self.volume_total_ml

class Estoque:
    def __init__(self, df_estoque):
        self.produtos = {}
        for _, row in df_estoque.iterrows():
            p = Produto(row['Produto'], row['Volume (ml)'], row['Preço Unitário (R$/ml)'])
            p.set_initial()
            self.produtos[p.nome] = p

    def atualizar(self, nome, quantidade):
        produto = self.produtos[nome]
        produto.usar(quantidade)
        if produto.volume_total_ml <= produto.nivel_critico_frac * produto.volume_total_ml_initial:
            st.warning(f"⚠️ Estoque crítico de '{nome}': {produto.volume_total_ml} ml restante.")

class Servico:
    def __init__(self, nome, df_servicos):
        self.nome = nome
        df = df_servicos[df_servicos['Serviço'] == nome]
        # pivot para dict de sujidade
        self.parametros = {}
        if not df.empty:
            self.parametros = {
                'Baixa': df['Sujidade Baixa (ml)'].values[0],
                'Média': df['Sujidade Média (ml)'].values[0],
                'Alta': df['Sujidade Alta (ml)'].values[0]
            }

    def estimativa_ml(self, nivel_sujidade):
        return self.parametros.get(nivel_sujidade, 0)

class Carro:
    def __init__(self, tamanho, sujidade):
        self.tamanho = tamanho
        self.sujidade = sujidade

class Lavagem:
    def __init__(self, carro, servico, estoque):
        self.carro = carro
        self.servico = servico
        self.estoque = estoque
        self.produtos_usados_reais = {}
        self.timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    def estimativa_custo(self):
        # usar único produto por serviço
        ml = self.servico.estimativa_ml(self.carro.sujidade)
        prod = self.estoque.produtos[self.servico.nome_produto]
        return ml * prod.preco_unitario_ml

    def registrar_uso(self, quantidade):
        nome = self.servico.nome_produto
        self.produtos_usados_reais[nome] = quantidade
        self.estoque.atualizar(nome, quantidade)

    def comparar(self):
        estimado = self.servico.estimativa_ml(self.carro.sujidade)
        real = self.produtos_usados_reais.get(self.servico.nome_produto, 0)
        custo_estimado = estimado * self.estoque.produtos[self.servico.nome_produto].preco_unitario_ml
        custo_real = real * self.estoque.produtos[self.servico.nome_produto].preco_unitario_ml
        return estimado, real, custo_estimado, custo_real

# Função para salvar histórico no workbook em memória
def salvar_historico(wb, lavagem):
    if 'Histórico de Lavagens' not in wb.sheetnames:
        ws = wb.create_sheet('Histórico de Lavagens')
        ws.append(['Timestamp', 'Serviço', 'Sujidade', 'Estimado (ml)', 'Real (ml)', 'Custo Estimado (R$)', 'Custo Real (R$)'])
    else:
        ws = wb['Histórico de Lavagens']
    est_ml, real_ml, est_c, real_c = lavagem.comparar()
    ws.append([lavagem.timestamp, lavagem.servico.nome, lavagem.carro.sujidade, est_ml, real_ml, est_c, real_c])

# -------------------------
# Streamlit UI
# -------------------------
st.title("Neutron Detail - Controle de Lavagens")

# Upload da planilha
uploaded = st.file_uploader("Envie sua planilha Excel", type=['xlsx'])
if uploaded:
    # ler planilha
dfs = load_excel = pd.read_excel(uploaded, sheet_name=None)
    df_estoque = dfs['Estoque']
    df_servicos = dfs['Serviços']

    # criar objetos
    estoque = Estoque(df_estoque)
   
    # Seleção de serviço e carro
    servico_nome = st.selectbox("Serviço", df_servicos['Serviço'].unique())
    sujidade = st.selectbox("Nível de Sujidade", ['Baixa', 'Média', 'Alta'])

    # instancia serviço com atributo produto associado
    servico = Servico(servico_nome, df_servicos)
    # mapear nome de produto usado = mesma coluna "Produto"
    servico.nome_produto = df_servicos[df_servicos['Serviço'] == servico_nome]['Produto'].values[0]

    carro = Carro(None, sujidade)

    lavagem = Lavagem(carro, servico, estoque)

    if st.button("Calcular Estimativa e Registrar Uso Real"):
        # estimativa
        est_ml = servico.estimativa_ml(sujidade)
        st.write(f"Estimativa de uso: {est_ml} ml")
        custo_est = est_ml * estoque.produtos[servico.nome_produto].preco_unitario_ml
        st.write(f"Custo estimado: R$ {custo_est:.2f}")

        # input real
        real_ml = st.number_input("Quantidade real usada (ml)", min_value=0, value=int(est_ml))
        lavagem.registrar_uso(real_ml)
        est_ml, real_ml, est_c, real_c = lavagem.comparar()
        st.success(f"Custo real: R$ {real_c:.2f}")

        # salvar no histórico e preparar download
        wb = load_workbook(uploaded)
        salvar_historico(wb, lavagem)
        # atualizar aba estoque
        ws2 = wb['Estoque']
        for idx, row in enumerate(df_estoque['Produto'], start=2):
            ws2.cell(row=idx, column=2, value=estoque.produtos[row].volume_total_ml)
        
        # gerar bytes para download
        stream = io.BytesIO()
        wb.save(stream)
        stream.seek(0)
        st.download_button("Baixar planilha atualizada", data=stream, file_name="Dados_atualizados.xlsx")
else:
    st.info("Faça upload da planilha modelo para começar.")
