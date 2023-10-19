from ._dependencias import *

class ExcelGrafico:
    def __init__(self, excel, arquivo, planilha, estilo, tipo):
        self.excel = excel
        self.arquivo = arquivo
        self.planilha = planilha
        self.tipo = tipo
        self.estilo = estilo
        self.grafico = self.planilha.Shapes.AddChart2(self. estilo, self.tipo)

    def DefinirFonteDados(self, intervalo):
        self.grafico.Chart.SetSourceData(intervalo)

    def MoverGrafico(self, celulaDestino):
        celula = self.planilha.Range(celulaDestino)
        self.grafico.Left = celula.Left
        self.grafico.Top = celula.Top

    def AlterarTitulo(self, titulo):
        self.grafico.Chart.ChartTitle.Text = titulo