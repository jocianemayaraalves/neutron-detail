import streamlit as st
import pandas as pd

# Funções para carregar e manipular o Excel
def load_excel(uploaded_file):
    dfs = pd.read_excel(uploaded_file, sheet_name=None)
    return dfs

# Função para calcular a estimativa de produtos e custos
def calcular_estimativa(servico, sujidade, produtos):
    estimativa = {}
    for produto in produtos:
        if sujidade == 'Baixa':
            estimativa[produto] = servico[produto]['Sujidade Baixa (ml)']
        elif sujidade == 'Média':
            estimativa[produto] = servico[produto]['Sujidade Média (ml)']
        else:
            estimativa[produto] = servico[produto]['Sujidade Alta (ml)']
    return estimativa

# Função para atualizar o estoque e histórico
def atualizar_estoque(estoque, produtos_usados):
    for produto, quantidade in produtos_usados.items():
        estoque[produto]['Volume (ml)'] -= quantidade
    return estoque

def salvar_historico(historico, dados_lavagem):
    historico.append(dados_lavagem)
    return historico

# Classe Produto
class Produto:
    def __init__(self, nome, volume, preco):
        self.nome = nome
        self.volume = volume
        self.preco = preco

# Classe Estoque
class Estoque:
    def __init__(self):
        self.produtos = {}

    def adicionar_produto(self, produto):
        self.produtos[produto.nome] = produto

    def obter_produto(self, nome_produto):
        return self.produtos.get(nome_produto, None)

# Classe Servico
class Servico:
    def __init__(self, nome, produtos):
        self.nome = nome
        self.produtos = produtos

# Função principal do app
def app():
    # Carregar planilha
    uploaded = st.file_uploader("Faça o upload da planilha Excel", type="xlsx")
    if uploaded is not None:
        dfs = load_excel(uploaded)
        estoque_df = dfs["Estoque"]
        servicos_df = dfs["Serviços"]

        # Criando os objetos de estoque
        estoque = Estoque()
        for index, row in estoque_df.iterrows():
            produto = Produto(row['Produto'], row['Volume (ml)'], row['Preço Unitário (R$/ml)'])
            estoque.adicionar_produto(produto)

        # Criando os serviços
        servicos = {}
        for index, row in servicos_df.iterrows():
            servico = Servico(row['Serviço'], {
                row['Produto']: {
                    'Sujidade Baixa (ml)': row['Sujidade Baixa (ml)'],
                    'Sujidade Média (ml)': row['Sujidade Média (ml)'],
                    'Sujidade Alta (ml)': row['Sujidade Alta (ml)'],
                }
            })
            servicos[row['Serviço']] = servico

        # Seleção no app
        servico_selecionado = st.selectbox("Selecione o serviço", list(servicos.keys()))
        sujidade_selecionada = st.selectbox("Selecione o nível de sujidade", ["Baixa", "Média", "Alta"])

        if st.button("Calcular Estimativa"):
            servico = servicos[servico_selecionado]
            estimativa = calcular_estimativa(servico.produtos, sujidade_selecionada, servico.produtos)
            st.write("Estimativa de produtos e custos:")
            for produto, quantidade in estimativa.items():
                st.write(f"{produto}: {quantidade} ml")

            # Entrada de dados reais
            produtos_usados = {}
            for produto in estimativa:
                quantidade_real = st.number_input(f"Qtd. Real usada de {produto} (ml):", min_value=0)
                produtos_usados[produto] = quantidade_real

            if st.button("Salvar e atualizar estoque"):
                estoque_atualizado = atualizar_estoque(estoque.produtos, produtos_usados)
                st.write("Estoque atualizado:")
                for produto, data in estoque_atualizado.items():
                    st.write(f"{produto}: {data['Volume (ml)']} ml restantes")

                # Salvar histórico
                historico = []
                dados_lavagem = {
                    "Serviço": servico_selecionado,
                    "Sujidade": sujidade_selecionada,
                    "Estimativa": estimativa,
                    "Produtos Usados": produtos_usados,
                    "Estoque Atualizado": estoque_atualizado
                }
                historico = salvar_historico(historico, dados_lavagem)
                st.write("Histórico de lavagens:")
                st.write(historico)

        # Option to download updated Excel
        if st.button("Baixar planilha atualizada"):
            with pd.ExcelWriter("Planilha_Atualizada_NeutronDetail.xlsx", engine="openpyxl") as writer:
                estoque_df.to_excel(writer, sheet_name="Estoque", index=False)
                servicos_df.to_excel(writer, sheet_name="Serviços", index=False)
                # Adicionar histórico em uma nova aba
                pd.DataFrame(historico).to_excel(writer, sheet_name="Histórico", index=False)
            st.write("Planilha atualizada foi gerada!")

if __name__ == "__main__":
    app()
