from ._dependencias import *
from .ExcelTabelaDinamica import *
from .ExcelGrafico import *

class ExcelPlanilha:
    def __init__(self, excel, arquivo, objeto):
        self.excel = excel
        self.arquivo = arquivo
        self.planilha = objeto
        self.nome = self.planilha.Name
        self.graficos = []
        self.tabelas = []

    def Visivel(self):
        self.planilha.Activate()

    def DefinirValorCelula(self, linha, coluna, valor):
        self.planilha.Cells(linha, coluna).Value = valor

    def ObterValorCelula(self, linha, coluna):
        return self.planilha.Cells(linha, coluna).Value

    def ObterEnderecoCelula(self, linha, coluna):
        return self.planilha.Cells(linha, coluna).GetAddress(False, False)

    def FormatarCelula(self, linha, coluna, font, size):
        self.planilha.Cells(linha, coluna).Font.Name = font
        self.planilha.Cells(linha, coluna).Font.Size = size

    def ObterFormatacaoCelula(self, linha, coluna):
        fonte = self.planilha.Cells(linha, coluna).Font.Name
        tamanho = self.planilha.Cells(linha, coluna).Font.Size
        return (fonte, tamanho)

    def InserirTabelaDinamica(self, planilhaDados, dados, nomeTabela, celula):
        tabela = ExcelTabelaDinamica(self.excel, self.arquivo, self.planilha, 
                                     planilhaDados, dados, nomeTabela, celula)
        self.tabelas.append(tabela)
        return tabela

    def InserirGrafico(self, estilo, tipo):
        grafico = ExcelGrafico(self.excel, self.arquivo,
                               self.planilha, estilo, tipo)
        self.graficos.append(grafico)
        return grafico

    def ObterQntdLinhas(self, coluna):
        qntdLinhas = self.planilha.Cells(
            self.planilha.Rows.Count, coluna).End(constants.xlUp).Row
        return qntdLinhas

    def ObterQntdColunas(self, linha):
        qntdColunas = self.planilha.Cells(
            linha, self.planilha.Columns.Count).End(constants.xlToLeft).Column
        return qntdColunas

    def AjustarColuna(self, coluna):
        self.planilha.Columns(coluna).AutoFit()