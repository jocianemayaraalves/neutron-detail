import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
import os

# -------------------------
# Classe Produto
# -------------------------
class Produto:
    def __init__(self, nome, volume_total_ml, preco_unitario_ml):
        self.nome = nome
        self.volume_total_ml = volume_total_ml
        self.preco_unitario_ml = preco_unitario_ml
        self.nivel_critico_frac = 0.05  # 5%

    def usar(self, quantidade_ml):
        if quantidade_ml > self.volume_total_ml:
            raise ValueError(f"Estoque insuficiente para '{self.nome}' ({quantidade_ml}ml solicitado, apenas {self.volume_total_ml}ml disponível)")
        self.volume_total_ml -= quantidade_ml

    def em_nivel_critico(self):
        return self.volume_total_ml <= (self.nivel_critico_frac * self.initial_volume_ml)

    def set_initial_volume(self):
        # registra volume inicial para cálculo de crítico
        self.initial_volume_ml = self.volume_total_ml


# -------------------------
# Classe Estoque
# -------------------------
class Estoque:
    def __init__(self):
        self.produtos = {}

    def adicionar_produto(self, produto: Produto):
        produto.set_initial_volume()
        self.produtos[produto.nome] = produto

    def atualizar(self, nome, usado_ml):
        if nome in self.produtos:
            p = self.produtos[nome]
            p.usar(usado_ml)
            if p.volume_total_ml <= p.nivel_critico_frac * p.initial_volume_ml:
                print(f"⚠️ Estoque crítico: '{nome}' restante {p.volume_total_ml}ml")
        else:
            raise KeyError(f"Produto '{nome}' não encontrado no estoque")


# -------------------------
# Classe Servico
# -------------------------
class Servico:
    def __init__(self, nome, parametros_df: pd.DataFrame):
        self.nome = nome
        # parametros_df tem colunas: Produto, Sujidade Baixa (ml), Sujidade Média (ml), Sujidade Alta (ml)
        # converter para dict: { 'Baixa': {prod:ml,...}, 'Média': {...}, 'Alta': {...} }
        levels = ['Baixa', 'Média', 'Alta']
        self.parametros = {}
        for lvl in levels:
            col = f"Sujidade {lvl} (ml)"
            # drop NaN
            df2 = parametros_df[['Produto', col]].dropna(subset=[col])
            self.parametros[lvl.lower()] = dict(zip(df2['Produto'], df2[col]))

    def estimar_uso(self, sujidade: str) -> dict:
        # retorna dict produto->ml estimado
        return self.parametros[sujidade.lower()]


# -------------------------
# Classe Carro
# -------------------------
class Carro:
    def __init__(self, tamanho: str, sujidade: str):
        self.tamanho = tamanho  # não usado na estimativa atual
        self.sujidade = sujidade


# -------------------------
# Classe Lavagem
# -------------------------
class Lavagem:
    def __init__(self, carro: Carro, servico: Servico, estoque: Estoque):
        self.carro = carro
        self.servico = servico
        self.estoque = estoque
        self.produtos_usados_reais = {}

    def estimativa_preco(self) -> float:
        uso_est = self.servico.estimar_uso(self.carro.sujidade)
        total = 0.0
        for nome, ml in uso_est.items():
            prod = self.estoque.produtos[nome]
            total += ml * prod.preco_unitario_ml
        return total

    def registrar_uso_real(self, nome: str, ml_real: float):
        self.produtos_usados_reais[nome] = ml_real
        self.estoque.atualizar(nome, ml_real)

    def calcular_preco_real(self) -> float:
        total = 0.0
        for nome, ml in self.produtos_usados_reais.items():
            prod = self.estoque.produtos[nome]
            total += ml * prod.preco_unitario_ml
        return total

    def comparar(self):
        est = self.estimativa_preco()
        real = self.calcular_preco_real()
        print(f"💡 Estimado: R$ {est:.2f} | Real: R$ {real:.2f}")


# -------------------------
# Funções de I/O Excel
# -------------------------
def load_data(filepath: str):
    """
    Lê o Excel e retorna (estoque, dict de serviços)
    """
    xls = pd.ExcelFile(filepath)
    # Estoque
    df_e = xls.parse('Estoque')
    estoque = Estoque()
    for _, row in df_e.iterrows():
        nome = row['Produto']
        vol = float(row['Volume (ml)'])
        preco = float(row['Preço Unitário (R$/ml)'])
        estoque.adicionar_produto(Produto(nome, vol, preco))
    # Serviços
    df_s = xls.parse('Serviços')
    servicos = {}
    for serv_name, grp in df_s.groupby('Serviço'):
        servicos[serv_name] = Servico(serv_name, grp)
    return estoque, servicos


def salvar_historico_excel(filepath: str, lavagem: Lavagem):
    """
    Adiciona registros de lavagem na aba 'Histórico de Lavagens'
    """
    if not os.path.exists(filepath):
        raise FileNotFoundError(f"Arquivo não encontrado: {filepath}")
    wb = load_workbook(filepath)
    if 'Histórico de Lavagens' not in wb.sheetnames:
        ws = wb.create_sheet('Histórico de Lavagens')
        ws.append(['Data/Hora', 'Tamanho', 'Sujidade', 'Serviço',
                   'Produto', 'Estimado (ml)', 'Real (ml)', 'Custo Estimado (R$)', 'Estoque Restante (ml)'])
    else:
        ws = wb['Histórico de Lavagens']

    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    est = lavagem.servico.estimar_uso(lavagem.carro.sujidade)
    custo_est = lavagem.estimativa_preco()
    for nome, real_ml in lavagem.produtos_usados_reais.items():
        estim_ml = est.get(nome, 0)
        restante = lavagem.estoque.produtos[nome].volume_total_ml
        ws.append([now,
                   lavagem.carro.tamanho,
                   lavagem.carro.sujidade,
                   lavagem.servico.nome,
                   nome,
                   estim_ml,
                   real_ml,
                   round(estim_ml * lavagem.estoque.produtos[nome].preco_unitario_ml, 2),
                   restante])
    wb.save(filepath)
    print("✅ Histórico salvo no Excel.")


# -------------------------
# Exemplo de uso
# -------------------------
if __name__ == '__main__':
    # caminho para sua planilha
    arquivo = 'Dados base.xlsx'
    estoque, servicos = load_data(arquivo)

    # Exemplo: realizar lavagem
    carro = Carro(tamanho='Pequeno', sujidade='Média')
    servico = servicos.get('Lavagem com Snow Foam')
    lavagem = Lavagem(carro, servico, estoque)

    print("Estimativa de custo: R$", round(lavagem.estimativa_preco(), 2))

    # registrar uso real (pode ser input ou valores de teste)
    for prod_nome, est_ml in servico.estimar_uso(carro.sujidade).items():
        real_ml = float(input(f"Informe ml real usado de {prod_nome}: "))
        lavagem.registrar_uso_real(prod_nome, real_ml)

    lavagem.comparar()
    salvar_historico_excel(arquivo, lavagem)
