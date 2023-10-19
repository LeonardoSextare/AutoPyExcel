from ._dependencias import *
from .ExcelArquivo import *


class ExcelAplicativo():
    def __init__(self):
        self.excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.arquivos = []
        self.DefinirPlanilhaPadrao(1)

    def AbrirArquivo(self, caminhoArquivo):
        arquivo = ExcelArquivo(self.excel, caminhoArquivo)
        self.arquivos.append(arquivo)
        return arquivo

    def Visivel(self):
        self.excel.Visible = 1

    def Esconder(self):
        self.excel.Visible = 0

    def Sair(self):
        for arquivo in self.arquivos:
            arquivo.Fechar()
        self.excel.Quit()

    def DefinirPlanilhaPadrao(self, numeroPlanilha):
        self.excel.SheetsInNewWorkbook = numeroPlanilha

    # def CriarArquivo(self, caminhoArquivo):
    #     print("Não implementado")
