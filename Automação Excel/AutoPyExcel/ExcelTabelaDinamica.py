from ._dependencias import *


class ExcelTabelaDinamica():
    def __init__(self, excel, arquivo, planilha, planilhaDados, dados, nome2, celula):
        self.excel = excel
        self.arquivo = arquivo
        self.planilha = planilha
        self.planilhaDados = planilhaDados
        self.dados = dados
        self.nome = nome2
        self.celula = celula
        self.linhas = {}
        self.colunas = {}
        self.valores = {}
        self.filtros = {}
        self.cache = self.arquivo.PivotCaches().Create(
            1, self.planilhaDados.Range(self.dados).CurrentRegion)
        self.tabela = self.cache.CreatePivotTable(
            planilha.Range(self.celula), self.nome)

    def AdicionarLinha(self, variavel):
        self.linhas[variavel] = self.tabela.PivotFields(variavel)
        self.linhas[variavel].Orientation = 1
        self.linhas[variavel].Position = 1

    def AdicionarValor(self, variavel, nome, calculo):
        self.valores[variavel] = self.tabela.PivotFields(variavel)
        self.valores[variavel].Orientation = 4   # 4 para Valor
        # -4112 para definir como Contagem
        self.valores[variavel].Function = -4112
        self.valores[variavel].Name = nome  # Nome da Coluna
        self.valores[variavel].Calculation = calculo
        if calculo == 5:
            self.valores[variavel].BaseField = variavel

    def AgruparValores(self, celula, inicio=True, fim=True, intervalo=True):
        grupo = self.planilha.Range(celula)
        grupo.Group(inicio, fim, intervalo)
