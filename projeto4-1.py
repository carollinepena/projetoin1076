from tkinter import *
from tkinter import filedialog as dlg
import csv
import matplotlib.pyplot as plt
from fpdf import FPDF
import pathlib
import os
from tkinter import filedialog
from tkinter.filedialog import asksaveasfile, askopenfilename

tituloGlobal:str = ""
descricaoGlobal:str = ""
caminhoDeArquivo:str = ""
ativar1:bool = False
ativar2:bool = False
dadosPlanilhaGlobal:dict = {}
listaCabecalhoGlobal:list = []
listaChavesGlobal:list = []
local = ""
cabecalhosAdicionados:list = []
cabecalhosSelecionados:dict = {}
listaDecabecalhosSelecionados = []
dadosComEtiqueta:dict = {}
testeDeErro:bool = False
dicioParaCabecalho:dict = {}
cabecalhoDeEtiqueta:str = ""
pdfPath = ""

class janelaPricipal:

    def __init__(self):

        self.janelaP = Tk()
        self.janelaP.title("Tela Principal")
        self.janelaP.geometry("600x400")



        self.fontePrincipal = ("Helvetica", "12","bold")
        self.fonteSecundaria = ("Helvetica", "12")

        # 1 container
        self.primeiroBloco = Frame(self.janelaP)
        self.primeiroBloco["pady"] = 20
        self.primeiroBloco.pack()

        self.titulo = Label(self.primeiroBloco, text="Organizer")
        self.titulo["font"] = ("Helvetica", "25", "bold")
        self.titulo.pack()

        # 2 container
        self.segundoBloco = Frame(self.janelaP)
        self.segundoBloco["padx"] = 20
        self.segundoBloco.pack()

        self.projetoLabel = Label(self.segundoBloco, text="Título:", font=self.fontePrincipal)
        self.projetoLabel["padx"] = 20
        self.projetoLabel.pack(side=LEFT)

        self.projeto = Entry(self.segundoBloco)
        self.projeto["width"] = 60
        self.projeto["font"] = self.fonteSecundaria
        self.projeto.pack(side=LEFT)

        # 3 container
        self.terceiroBloco = Frame(self.janelaP)
        self.terceiroBloco["padx"] = 20
        self.terceiroBloco.pack()

        self.descricaoLabel = Label(self.terceiroBloco, text="Descrição: ", font=self.fontePrincipal)
        self.descricaoLabel.pack(side=LEFT)
        self.descricao = Entry(self.terceiroBloco)
        self.descricao["width"] = 60
        self.descricao["font"] = self.fonteSecundaria
        self.descricao.pack(side=LEFT)

        # 4 container
        self.quartoBloco = Frame(self.janelaP)
        self.quartoBloco["pady"] = 30

        self.quartoBloco.pack()

        self.salvar = Button(self.quartoBloco)
        self.salvar["text"] = "SALVAR"
        self.salvar["font"] = ("Helvetica", "8")
        self.salvar["width"] = 12
        self.salvar["command"] = self.salvarArgumentos
        self.salvar.pack()

        self.mensagem = Label(self.quartoBloco, text=" ",  foreground="red4")
        self.mensagem["font"] = ("Heveltica", "10")
        self.mensagem.pack()

        # 5 container
        self.quintoBloco = Frame(self.janelaP)

        self.quintoBloco.pack()

        self.instrucoesDeUso = Label(self.quintoBloco, text=' Instruções gerais para uso:\n'
                                                            ' \n'
                                                            '  1) Você precisa inserir o Título e a Descricão do relatório; \n'
                                                            '  2) Você vai precisar inserir os cabeçalhos da mesma forma que estão descritos na planilha; \n'
                                                            '  3) Você poderá selecionar os cabeçalhos para compor o relatório; \n'
                                                            '  4) Escolha se deseja gerar um arquivo txt ou pdf;\n ',
                                     justify=LEFT,
                                     bd=1,
                                     relief=GROOVE,
                                     bg="gray92")
        self.instrucoesDeUso["width"] = 70
        self.instrucoesDeUso.pack()

        self.janelaP.mainloop()


    def  salvarArgumentos(self):
        t1 = self.projeto.get()
        d2 = self.descricao.get()
        global tituloGlobal
        global descricaoGlobal

        if t1 != "" and d2 != "":

            tituloGlobal = t1
            descricaoGlobal = d2
            self.janelaP.destroy()
            janelaDeCabecalho()

        else:
            self.mensagem["text"] = "Insira o Título e/ou a Descrição do Relatório"
            self.mensagem["font"] = ("Heveltica", "10")


class janelaDeCabecalho():

    def __init__(self):

        self.janela3 = Tk()
        self.janela3.title("Definir Cabeçalho")
        self.janela3.geometry("600x650")


        # 1 container
        self.fontePrincipal = ("Helvetica", "12","bold")
        self.primeiroBloco3 = Frame(self.janela3)
        self.primeiroBloco3["pady"] = 10
        self.primeiroBloco3.pack()

        self.nomeProjeto = Label(self.primeiroBloco3, text=tituloGlobal, font= self.fontePrincipal)
        self.nomeProjeto.pack(side=TOP)

        # 2 container
        self.segundoBloco3 = Frame(self.janela3)
        self.segundoBloco3["pady"] = 10
        self.segundoBloco3.pack()

        self.descricaoRelatorio = Label(self.segundoBloco3, text=descricaoGlobal)
        self.descricaoRelatorio["font"]= ("Helvetica", "10","bold")
        self.descricaoRelatorio.pack(side=RIGHT)

        self.tipoRelatorio = Label(self.segundoBloco3, text= "Descrição:")
        self.tipoRelatorio["font"]=("Helvetica", "10","bold")
        self.tipoRelatorio.pack(side = LEFT)

        # 3 container

        self.terceiroBloco3 = Frame(self.janela3)
        self.terceiroBloco3["padx"] = 20
        self.terceiroBloco3.pack()

        self.corpoProjeto = Label(self.terceiroBloco3)
        self.corpoProjeto["pady"] = 20
        self.corpoProjeto["text"] = "Insira os itens do cabeçalho de sua planilha"
        self.corpoProjeto.pack(side=TOP)

        # 4 container

        self.quartoBloco3 = Frame(self.janela3)
        self.quartoBloco3["padx"] = 10
        self.quartoBloco3.pack()

        self.itemCab = Entry(self.quartoBloco3)
        self.itemCab["width"] = 50
        self.itemCab.pack(side=LEFT)

        self.salvarItem = Button(self.quartoBloco3)
        self.salvarItem["text"] = "Inserir"
        self.salvarItem["font"] = ("Heveltica", "10")
        self.salvarItem["width"] = 12

        self.salvarItem["command"] = self.pegarItem
        self.salvarItem.pack(side=RIGHT)

        # 5 container

        self.quintoBloco3 = Frame(self.janela3)

        self.quintoBloco3["pady"] = 10
        self.quintoBloco3.pack(side=TOP)

        self.enfimLista = Listbox(self.quintoBloco3)
        self.enfimLista.pack(side=TOP,fill=BOTH, expand=False)
        self.enfimLista["width"] = 68
        self.enfimLista["height"] = 20

        self.rotuloItem = StringVar()

        self.enfimLista.bind("<<ListboxSelect>>")

        self.barraScroll = Scrollbar(self.quintoBloco3)

        self.barraScroll.pack(side=RIGHT, fill=Y)
        self.enfimLista.pack(side=LEFT, fill=Y)

        self.enfimLista.configure(yscrollcommand=self.barraScroll.set)
        self.barraScroll.configure(command=self.enfimLista.yview)


        # 6 container
        self.sextoBloco3 = Frame(self.janela3)
        self.sextoBloco3["pady"] = 30
        self.sextoBloco3.pack()

        self.proximaEtapa = Button(self.sextoBloco3)
        self.proximaEtapa["text"] = "Proximo"
        self.proximaEtapa["width"] = 12
        self.proximaEtapa["font"] = ("Heveltica", "10")

        self.proximaEtapa["command"] = self.chamarSecundaria
        self.proximaEtapa.pack(side=RIGHT)

        self.botaoRemover = Button(self.sextoBloco3)
        self.botaoRemover["text"] = "Voltar"
        self.botaoRemover["width"] = 12
        self.botaoRemover["font"] = ("Heveltica", "10")
        self.botaoRemover["command"] = self.voltarPrincipal
        self.botaoRemover.pack(side=LEFT)

        self.botaoPop = Button(self.sextoBloco3)
        self.botaoPop["text"] = "Remover"
        self.botaoPop["width"] = 12
        self.botaoPop["font"] = ("Heveltica", "10")
        self.botaoPop["command"] = self.popItem
        self.botaoPop.pack(side=LEFT)

        self.janela3.mainloop()


    def pegarItem(self):
        global cabecalhosAdicionados
        itens = self.itemCab.get()
        listaBox = []
        var = StringVar(value=listaBox)
        if itens:
            cabecalhosAdicionados.append(itens)
            self.enfimLista.insert(END, itens)


    def popItem(self):
        global cabecalhosAdicionados

        try:
            itemList = self.enfimLista.curselection()[0]
            nomeDoItem = self.enfimLista.get(itemList)
            self.rotuloItem.set(nomeDoItem)
            #print("teste")
            if nomeDoItem in cabecalhosAdicionados:


                    cabecalhosAdicionados.remove(nomeDoItem)
                    self.enfimLista.delete(itemList)
        except IndexError:
            pass



    def voltarPrincipal(self):
        global cabecalhosAdicionados


        self.janela3.destroy()
        cabecalhosAdicionados = []
        janelaPricipal()


    def chamarSecundaria(self):

        janelaSecundaria()
        self.janela3.destroy()

class janelaSecundaria():

    def __init__(self):
        self.janela2 = Tk()
        self.janela2.title("Selecionar arquivo")
        self.janela2.geometry("600x200")

        #1 container
        self.primeiroBloco2 = Frame(self.janela2)
        self.primeiroBloco2["pady"] = 20
        self.primeiroBloco2.pack()

        self.nomeItem = Label(self.primeiroBloco2)
        self.nomeItem["text"] = "Insira sua planilha"
        self.nomeItem["font"] = ("Heveltica", "14")
        self.nomeItem.pack(side=TOP)

        # 3 container
        self.terceiroBloco2 = Frame(self.janela2)
        self.terceiroBloco2["padx"] = 60
        self.terceiroBloco2["width"] = 60
        self.terceiroBloco2.pack()
        self.descricaoLabel = Label(self.terceiroBloco2, text="")
        self.descricaoLabel.pack(side=LEFT)

        self.nomeArquivo = Label(self.terceiroBloco2)
        self.nomeArquivo["text"] = "Arquivo.csv"
        self.nomeArquivo["width"] = 70
        self.nomeArquivo["relief"]= GROOVE
        self.nomeArquivo.configure(bg="gray92")
        self.nomeArquivo.pack()

        # 5 container
        self.quintoBloco2 = Frame(self.janela2)
        self.quintoBloco2["pady"] = 5
        self.quintoBloco2.pack()
        self.testelabele = Label(self.quintoBloco2)
        self.testelabele["text"] = " "
        self.testelabele.pack()

        # 2 container
        self.segundoBloco2 = Frame(self.janela2)
        self.segundoBloco2["padx"] = 30
        self.segundoBloco2.pack()

        self.adicionarItens = Button(self.segundoBloco2)
        self.adicionarItens["text"] = ("Selecionar")
        self.adicionarItens["font"] = ("Heveltica", "10")
        self.adicionarItens["width"] = 12
        self.adicionarItens["command"] = self.chamarJanelaSelecao
        self.adicionarItens.pack(side=LEFT)

        self.quartoBloco2 = Frame(self.janela2)
        self.quartoBloco2["padx"] = 180
        self.quartoBloco2.pack()

        self.proximaEtapa = Button(self.segundoBloco2)
        self.proximaEtapa["text"] = ("Seguinte")
        self.proximaEtapa["font"] = ("Heveltica", "10")
        self.proximaEtapa["width"] = 12
        self.proximaEtapa.pack(side=RIGHT)

    def chamarJanelaSelecao(self):  # ABRIR PASTA PARA INSERIR ARQUIV

        global caminhoDeArquivo
        global  testeDeErro
        global pdfPath

        Tk().withdraw()


        local = askopenfilename(title="Abrir arquivo.csv")
        caminhoDeArquivo = local

        csvPath = pathlib.Path(local)
        pdfPath = csvPath.parent

        if local:

            self.nomeArquivo["text"] = local
            self.nomeArquivo.pack(side=TOP)
            self.proximaEtapa["command"] = self.processar
            self.janela2.mainloop()

    def processar(self):  # FECHAR ESSA JANELA

        self.janela2.destroy()


        global ativar
        global dadosPlanilhaGlobal
        global listaCabecalhoGlobal
        global caminhoDeArquivo
        global local
        global testeDeErro

        dadosPlanilha= {}
        listaCabecalho = []

        ativar = True

        try:
            if ativar == True:

                with open (caminhoDeArquivo) as arquivocsv:

                    planilha = csv.DictReader(arquivocsv, delimiter=",")

                    for linha in planilha:
                        for chave, valor in linha.items():
                            if chave not in dadosPlanilha:
                                dadosPlanilha[chave] = []
                            dadosPlanilha[chave].append(valor)

                    dadosPlanilhaGlobal = dadosPlanilha

                for codigo in dadosPlanilha:
                    listaCabecalho.append(codigo)

                listaCabecalhoGlobal = listaCabecalho

            #print(dadosPlanilhaGlobal)
            #janelaDeCheckBook()
            janelaDeEtiquetas()


        except (UnicodeDecodeError, FileNotFoundError):
            local = ""
            chamarJanelaDeErro()


class chamarJanelaDeErro:
    def __init__(self):

            self.janeladeErro = Tk()
            self.janeladeErro.geometry("700x100")
            self.janeladeErro.title("ERRO")
            self.janeladeErro.attributes("-topmost", True)

            #1 container
            self.blocoDoErrinho = Frame(self.janeladeErro)
            self.blocoDoErrinho["pady"] = 20
            self.blocoDoErrinho.pack()

            self.mensagemErro = Label(self.blocoDoErrinho)
            self.mensagemErro["text"] = "Ocorreu um problema ao ler seu arquivo, ou ele está em um formato não aceito. Volte para janela de seleção"
            self.mensagemErro.pack()

            # 2 container
            self.botaoDoerro = Button(self.blocoDoErrinho)
            self.botaoDoerro["text"] = "VOLTAR"
            self.botaoDoerro["width"] = 12
            self.botaoDoerro["command"] = self.voltarPraSelecao
            self.botaoDoerro.pack()



    def voltarPraSelecao(self):
        global testeDeErro
        testeDeErro = False
        janelaSecundaria()
        self.janeladeErro.destroy()

class janelaDeEtiquetas:
    def __init__(self):
        self.janelaEtiqueta = Tk()
        self.janelaEtiqueta.title("Selecionar Partes")
        self.janelaEtiqueta.geometry("600x660")

        # 1 container
        self.primeiroEtiqueta = Frame(self.janelaEtiqueta)
        self.primeiroEtiqueta["padx"] = 20
        self.primeiroEtiqueta["pady"] = 20
        self.primeiroEtiqueta.pack()

        self.perguntaEtiqueta = Label(self.primeiroEtiqueta)
        self.perguntaEtiqueta["text"] = "Algum desses itens é uma etiqueta?"
        self.perguntaEtiqueta["font"] = ("Heveltica", "14")
        self.perguntaEtiqueta.pack()

        # 2 container
        self.segundoEtiqueta = Frame(self.janelaEtiqueta)
        self.segundoEtiqueta["width"] = 10
        self.segundoEtiqueta["height"] = 30

        self.segundoEtiqueta.pack(side=TOP)

        dicioParaCabecalho = cabecalhosAdicionados
        #print(cabecalhosAdicionados)


        # 3 container

        self.testeLista = Frame(self.janelaEtiqueta)
        self.testeLista.pack(side=TOP)

        self.listaEtiqueta = Listbox(self.testeLista)
        self.listaEtiqueta["width"] = 40
        self.listaEtiqueta["height"] = 20
        self.listaEtiqueta.pack(side=LEFT, fill=BOTH, expand=False)
        self.rotuloEtiqueta = StringVar()
        self.listaEtiqueta.bind("<<ListboxSelect>>")

        self.barraScroll4 = Scrollbar(self.testeLista)
        self.barraScroll4.pack(side=RIGHT, fill=Y)
        self.listaEtiqueta.pack(side=LEFT, fill=Y)

        self.listaEtiqueta.configure(yscrollcommand=self.barraScroll4.set)
        self.barraScroll4.configure(command=self.listaEtiqueta.yview)

        for u in range (0, len(cabecalhosAdicionados)):
            self.listaEtiqueta.insert(END, cabecalhosAdicionados[u])


        self.terceiroEtiqueta = Frame(self.janelaEtiqueta)
        self.terceiroEtiqueta["padx"] = 20
        self.terceiroEtiqueta["pady"] = 20
        self.terceiroEtiqueta.pack()

        self.botaoSeguir = Button(self.terceiroEtiqueta)
        self.botaoSeguir["text"] = "SIM"
        self.botaoSeguir["width"] = 12
        self.botaoSeguir["command"] = self.pegaEtiqueta
        self.botaoSeguir.pack(side=RIGHT)

        self.botaoCancelar = Button(self.terceiroEtiqueta)
        self.botaoCancelar["text"] = "NÃO"
        self.botaoCancelar["width"] = 12
        self.botaoCancelar["command"] = self.chamarCheckSemEtiqueta
        self.botaoCancelar.pack(side=LEFT)

        # 4 container

        self.quartoEtiqueta = Frame(self.janelaEtiqueta)

        self.quartoEtiqueta.pack()

        self.instrucoesDeEtiqueta = Label(self.quartoEtiqueta, text='\n Seleção de Etiqueta:\n'
                                                            ' \n'
                                                            '  1) Selecione uma etiqueta caso exista uma coluna de etiquetas em sua planila; \n'
                                                            '  2) Você só pode adicionar uma etiqueta; \n'
                                                            '  3) Caso sua planilha não possua, clique em "Não" ; \n'
                                                            ,
                                     justify=LEFT,
                                     bd=1,
                                     relief=GROOVE,
                                     bg="gray92")

        self.instrucoesDeEtiqueta["width"] = 70
        self.instrucoesDeEtiqueta.pack()


    def pegaEtiqueta (self):

        global cabecalhosAdicionados
        global cabecalhoDeEtiqueta
        try:
            posicaoEtiqueta = self.listaEtiqueta.curselection()[0]
            aEtiqueta = self.listaEtiqueta.get(posicaoEtiqueta)

            cabecalhoDeEtiqueta = aEtiqueta
            #print(aEtiqueta)

            self.janelaEtiqueta.destroy()
            janelaDeCheckBook()

        except IndexError:
            pass



    def chamarCheckSemEtiqueta(self):
        global cabecalhoDeEtiqueta
        cabecalhoDeEtiqueta = "1/*-/123434543"
        self.janelaEtiqueta.destroy()
        janelaDeCheckBook()


class janelaDeCheckBook:

    def __init__(self):
        self.janelaCheck = Tk()
        self.janelaCheck.title("Selecionar Partes")
        self.janelaCheck.geometry("600x700")


        # 1 container

        self.primeiroBloco4 = Frame(self.janelaCheck)
        self.primeiroBloco4["padx"] = 20
        self.primeiroBloco4["pady"] = 20
        self.primeiroBloco4.pack()

        self.selecioneCab = Label(self.primeiroBloco4, text=local)
        self.selecioneCab["text"] = "Selecione os cabeçalhos para formar seu relatório"
        self.selecioneCab["font"] = ("Heveltica", "14")
        self.selecioneCab.pack()

        # 2 container

        global cabecalhosAdicionados
        global listaDecabecalhosSelecionados

        self.segundoBloco4 = Frame(self.janelaCheck)
        self.segundoBloco4["padx"] = 20
        self.segundoBloco4.pack()

        listaCheck = cabecalhosAdicionados

        # 3 container

        self.terceiroBloco4 = Frame(self.janelaCheck)
        self.terceiroBloco4.pack()

        self.labelInserir = Label(self.terceiroBloco4)

        # 6 container

        self.sextoBloco5 = Frame(self.janelaCheck)
        self.sextoBloco5.pack()

        self.cabDispo = Label(self.sextoBloco5)
        self.cabDispo["padx"] = 60
        self.cabDispo["text"] = "Cabeçalhos Disponíveis"
        self.cabDispo["font"] = ("Heveltica", "10")
        self.cabDispo.pack(side=LEFT)

        self.cabeSelect = Label(self.sextoBloco5)
        self.cabeSelect["padx"] = 60
        self.cabeSelect["text"] = "Cabeçalhos Selecionados"
        self.cabeSelect["font"] = ("Heveltica", "10")
        self.cabeSelect.pack(side=RIGHT)




        # container espcial
        self.frame1 = Frame(self.janelaCheck)
        self.frame1["padx"] = 20
        self.frame1["pady"] = 10
        self.frame1.pack()

        self.enfimLista1 = Listbox(self.frame1)

        self.enfimLista1["width"] = 40
        self.enfimLista1["height"] = 20
        self.enfimLista1.pack(side=LEFT, fill=BOTH, expand=False)
        self.rotuloItem1 = StringVar()
        self.enfimLista1.bind("<<ListboxSelect>>", self.inserirIntem1)

        self.enfimLista2 = Listbox(self.frame1)
        self.enfimLista2["width"] = 40
        self.enfimLista2["height"] = 20
        self.enfimLista2.pack(side=RIGHT, fill=BOTH, expand=False)
        self.rotuloItem2 = StringVar()
        self.enfimLista2.bind("<<ListboxSelect>>", self.removerIntem2)

        for v in range(0, len(cabecalhosAdicionados)):
            self.enfimLista1.insert(END, cabecalhosAdicionados[v])

        #scrolls
        self.barraScroll1 = Scrollbar(self.frame1)
        self.barraScroll1.pack(side=LEFT, fill=Y)
        self.enfimLista1.configure(yscrollcommand=self.barraScroll1.set)
        self.barraScroll1.configure(command=self.enfimLista1.yview)

        self.barraScroll2 = Scrollbar(self.frame1)
        self.barraScroll2.pack(side=RIGHT, fill=Y)
        self.enfimLista2.configure(yscrollcommand=self.barraScroll2.set)
        self.barraScroll2.configure(command=self.enfimLista2.yview)

        # 3 container
        self.terceiroBloco4 = Frame(self.janelaCheck)
        self.terceiroBloco4["padx"] = 20
        self.terceiroBloco4["pady"] = 30
        self.terceiroBloco4.pack()

        self.gerarRela = Button(self.terceiroBloco4)
        self.gerarRela["text"] = "Gerar Relatório TXT"
        self.gerarRela["font"] = ("Heveltica", "10")
        self.gerarRela["command"] = self.geracaoRelatorio
        self.gerarRela.pack(side=LEFT)

        self.gerarPDF = Button(self.terceiroBloco4)
        self.gerarPDF["text"] = "Gerar Relatório PDF"
        self.gerarPDF["font"] = ("Heveltica", "10")
        self.gerarPDF["command"] = self.geracaoPDF
        self.gerarPDF.pack(side=RIGHT)

        # 4 container
        self.quartoBloco4 = Frame(self.janelaCheck)
        self.quartoBloco4.pack()

        self.botaoImagem = Button(self.quartoBloco4)
        self.botaoImagem["text"] = "Gerar imagem"
        self.botaoImagem["command"] = self.contarNumeros
        self.botaoImagem["width"] = 12
        self.botaoImagem.pack()

        # 5 container
        self.quintoBloco4 = Frame(self.janelaCheck)
        self.quintoBloco4["pady"] = 10
        self.quintoBloco4.pack(side=BOTTOM)

        self.instrucoesDeUso = Label(self.quintoBloco4, text=' Instruções para gerar relatório:\n'
                                                            ' \n'
                                                            '  1) Se você não tiver inserido cabeçalhos, o relatório criado estara em branco; \n'
                                                            '  2) Você pode inserir os cabeçalhos da esquerda para a direita; \n'
                                                            '  3) Você poderá selecionar os cabeçalhos que deseja compor o relatório; \n'
                                                            '  4) Escolha se deseja gerar um arquivo txt ou pdf;\n ',
                                     justify=LEFT,
                                     bd=1,
                                     relief=GROOVE,
                                     bg="gray92")
        self.instrucoesDeUso["width"] = 70
        self.instrucoesDeUso.pack()

        self.janelaCheck.mainloop()



    def inserirIntem1(self, event):
        global cabecalhosAdicionados
        try:
            item1 = self.enfimLista1.curselection()[0]
            nomeDoItem1 = self.enfimLista1.get(item1)
            self.rotuloItem1.set(nomeDoItem1)

            if nomeDoItem1 in cabecalhosAdicionados:


                listaDecabecalhosSelecionados.append(nomeDoItem1)

                self.enfimLista1.delete(item1)
                self.enfimLista2.insert(END, nomeDoItem1)

        except IndexError:
            pass

        #print(cabecalhosAdicionados)


    def removerIntem2(self, event):
        global cabecalhosAdicionados
        global listaDecabecalhosSelecionados

        try:
            item2 = self.enfimLista2.curselection()[0]
            nomeDoItem2 = self.enfimLista2.get(item2)
            self.enfimLista2.delete(item2)
            self.enfimLista1.insert(END,nomeDoItem2)
            self.rotuloItem2.set(nomeDoItem2)
            listaDecabecalhosSelecionados.remove(nomeDoItem2)




            if nomeDoItem2 in listaDecabecalhosSelecionados:



                listaDecabecalhosSelecionados.remove(nomeDoItem2)

                self.enfimLista2.delete(item2)
                self.enfimLista1.insert(END, nomeDoItem2)



        except (IndexError,ValueError ):
            pass



            self.var = 1



    def geracaoRelatorio(self):
        salvo = False
        global cabecalhoDeEtiqueta
        global cabecalhosSelecionados
        global listaDecabecalhosSelecionados


        cabecalhosSelecionados= {}

        etiqueta = (str(cabecalhoDeEtiqueta))



        Tk().withdraw()

        for q in range (0,len(listaDecabecalhosSelecionados)):
            cabecalhosSelecionados[listaDecabecalhosSelecionados[q]] = 1


        arquivo = filedialog.asksaveasfile(mode='w',initialfile = "Relatorio txt - 01", defaultextension=".txt")
        try:
            arquivo.write(tituloGlobal+ "\n")
            arquivo.write(descricaoGlobal + "\n")

            for key in cabecalhosSelecionados:
                concatena = ""
                if cabecalhosSelecionados[key] == 1:

                    if key in dadosPlanilhaGlobal:

                        if etiqueta in dadosPlanilhaGlobal:
                            if cabecalhoDeEtiqueta in dadosPlanilhaGlobal:
                                valoresDaKey = dadosPlanilhaGlobal[key]

                                valoresEtiqueta = dadosPlanilhaGlobal[cabecalhoDeEtiqueta]

                                for t in range (0,len(valoresDaKey)):
                                    valoresDaKey[t] = valoresDaKey[t]
                                    concatena = concatena + valoresEtiqueta[t] +":"+ (valoresDaKey[t]).replace("\n", " ") + "\n"
                                arquivo.write(key + ":" + "\n")
                                arquivo.write(concatena)
                                arquivo.write("\n")
                                arquivo.write("\n")
                                salvo = TRUE


                        else:

                            valoresDaKey = dadosPlanilhaGlobal[key]

                            for z in range(0, len(valoresDaKey)):
                                ordem = str(z+1)
                                concatena = concatena  + ordem + ": " + (valoresDaKey[z]).replace("\n", " ") + "\n"
                            arquivo.write(key + ":  " + "\n")
                            arquivo.write(concatena)
                            arquivo.write("\n")
                            arquivo.write("\n")
                            salvo = TRUE




                else:
                    pass

            arquivo.close()

        except AttributeError:
            pass


        try:
            if salvo == True:
                self.janelaDeSalvo = Tk()
                self.janelaDeSalvo.title("Arquivo Salvo")
                self.janelaDeSalvo.geometry("200x100")
                self.blocoSalvo = Frame(self.janelaDeSalvo)
                self.blocoSalvo["pady"] = 20
                self.blocoSalvo.pack()

                self.mensagemSalvo = Label(self.blocoSalvo)
                self.mensagemSalvo["text"] = "Arquivo Salvo"
                self.mensagemSalvo.pack()

                self.botaoFecharSalvo = Button(self.blocoSalvo)
                self.botaoFecharSalvo["text"] = "ok"
                self.botaoFecharSalvo["width"] = 12
                self.botaoFecharSalvo["command"] = self.destruirJanelaDeSalvo
                self.botaoFecharSalvo.pack()

                arquivo = ""

                salvo = False

        except:
            pass



    def geracaoPDF(self):



        salvo = False
        global cabecalhoDeEtiqueta
        global cabecalhosSelecionados
        global pdfPath


        etiqueta = (str(cabecalhoDeEtiqueta))



        Tk().withdraw()

        for q in range(0, len(listaDecabecalhosSelecionados)):
            cabecalhosSelecionados[listaDecabecalhosSelecionados[q]] = 1

        arquivoT = open("relatorioT.txt", "w")

        try:


            for key in cabecalhosSelecionados:
                concatena = ""
                if cabecalhosSelecionados[key] == 1:

                    if key in dadosPlanilhaGlobal:
                        if etiqueta in dadosPlanilhaGlobal:
                            if cabecalhoDeEtiqueta in dadosPlanilhaGlobal:
                                valoresDaKey = dadosPlanilhaGlobal[key]
                                valoresEtiqueta = dadosPlanilhaGlobal[cabecalhoDeEtiqueta]
                                #print(valoresDaKey)
                                for k in range(0, len(valoresDaKey)):
                                    concatena = concatena + valoresEtiqueta[k] + ":" + (valoresDaKey[k]).replace("\n", " ") + "\n"
                                arquivoT.write(key + ":" + "\n")
                                arquivoT.write(concatena)
                                arquivoT.write("\n")
                                arquivoT.write("\n")
                                salvo = TRUE


                        else:

                            valoresDaKey = dadosPlanilhaGlobal[key]
                            #print(valoresDaKey)
                            #print("teste sem etiqueta")
                            for p in range(0, len(valoresDaKey)):
                                ordem = str(p + 1)
                                concatena = concatena + ordem + ": " + (valoresDaKey[p]).replace("\n", " ") + "\n"
                            arquivoT.write(key + ":  " + "\n")
                            arquivoT.write(concatena)
                            arquivoT.write("\n")
                            arquivoT.write("\n")
                            salvo = TRUE




                else:
                    pass

            arquivoT.close()
        except AttributeError:
            pass

        try:

            if salvo == True:
                self.janelaDeSalvo = Tk()
                self.janelaDeSalvo.title("Arquivo Salvo")
                self.janelaDeSalvo.geometry("400x100")
                self.blocoSalvo = Frame(self.janelaDeSalvo)
                self.blocoSalvo["pady"] = 20
                self.blocoSalvo.pack()

                self.mensagemSalvo = Label(self.blocoSalvo)
                self.mensagemSalvo["text"] = "Arquivo Salvo \n Procure por Relatório - PDF e renomeie o arquivo antes de gerar outro"
                self.mensagemSalvo.pack()

                self.botaoFecharSalvo = Button(self.blocoSalvo)
                self.botaoFecharSalvo["text"] = "ok"
                self.botaoFecharSalvo["width"] = 12
                self.botaoFecharSalvo["command"] = self.destruirJanelaDeSalvo
                self.botaoFecharSalvo.pack()

                arquivo = ""



            arquivoT.close()

        except:
            pass

        try:
            if arquivoT:

                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Arial", size=12)
                pdf.set_margins(10, 10 ,10)
                f = open("relatorioT.txt", "r")
                pdf.cell(200, 10, txt=tituloGlobal , ln=1, align="C")
                pdf.cell(200, 10, txt=descricaoGlobal, ln=1, align="C")
                for x in f:
                    pdf.multi_cell(180, 10, txt=x, align="J")

                #a = pdf.output(pdfPath + "\TesteProjeto.pdf", "F")


                a = pdf.output("Relatorio - PDF.pdf")

                f.close()
                os.remove("relatorioT.txt")

        except :
            pass


    def contarNumeros(self):  ###################################  CODIGO PARA GRAFICOS

        #global dadosPlanilhaGlobal
        #global cabecalhosSelecionados
        try:
            graficoDadosPlanilha: dict = dadosPlanilhaGlobal
            graficoCabecalhoSelecionado: list = listaDecabecalhosSelecionados
            cabecalhoDeGrafico: str = ""
            dadosParaGrafico: list = []

            if len(graficoCabecalhoSelecionado) == 1:
                cabecalhoDeGrafico = listaDecabecalhosSelecionados[0]

                for chave in graficoDadosPlanilha:
                    if cabecalhoDeGrafico == chave:
                        dadosParaGrafico = graficoDadosPlanilha[cabecalhoDeGrafico]



            numeros = dadosParaGrafico

            valorChaves = []
            valorOcorrencia = []


            dicioNumeros = {}

            for n in numeros:
                valor = dicioNumeros.get(n, -1)
                if valor > -1:  #
                    dicioNumeros[n] = valor + 1
                else:
                    dicioNumeros[n] = 1



            for chave in dicioNumeros:
                valorChaves.append(chave)

            for chave in dicioNumeros:
                valorOcorrencia.append(dicioNumeros[chave])


            tituloGrafico = ("Frequência: " + cabecalhoDeGrafico)
            plt.rcParams.update({'font.size': 20})



            plt.figure(figsize=(7, 7))
            plt.pie(x=valorOcorrencia, labels=valorChaves, autopct="%1.1f%%")
            plt.title(tituloGrafico)
            plt.show()
            plt.savefig("Figura - 1.pdf")


        except:
            pass


        #plt.xlabel(valorOcorrencia)
        #plt.ylabel(valorChaves)
        #plt.plot(valorChaves, valorOcorrencia)
        #plt.show()

        self.janelaCheck.mainloop()

    def destruirJanelaDeSalvo(self):
        self.janelaDeSalvo.destroy()



janelaPricipal()