from ._dependencias import *
from .ExcelPlanilha import *


class ExcelArquivo:
    def __init__(self, excel, caminhoArquivo):
        self.caminhoArquivo = caminhoArquivo
        self.excel = excel
        self.arquivo = self.excel.Workbooks.Open(caminhoArquivo)
        self.planilhas = {}

    def AdicionarPlanilha(self, nome):
        objeto = self.arquivo.Worksheets.Add()
        objeto.Name = nome
        planilha = ExcelPlanilha(self.excel, self.arquivo, objeto)
        self.planilhas[nome] = planilha
        return planilha

    def SelecionarPlanilha(self, planilha):
        objeto = self.arquivo.Sheets(planilha)
        planilha = ExcelPlanilha(self.excel, self.arquivo, objeto)
        self.planilhas[planilha.nome] = planilha
        return planilha

    def Salvar(self):
        self.arquivo.Save()

    def Fechar(self):
        self.planilhas = {}
        self.arquivo.Close()

    def DefinirAutor(self, autor):
        self.arquivo.Author = autor
