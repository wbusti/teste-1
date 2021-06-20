import tkinter as tk
from tkinter import ttk
import pandas as pd


class PrincipalRAD:
    def __init__(self, win):

#                componentes

        self.lblNome=tk.Label(win, text='Nome do Aluno: ')
        self.lblNota1=tk.Label(win, text='Nota 1')
        self.lblNota2=tk.Label(win, text='Nota 2')
        self.lblMedia=tk.Label(win, text='Media')
        self.txtNome=tk.Entry(bd=3)
        self.txtNota1=tk.Entry()
        self.txtNota2=tk.Entry()
        self.btnCalcular=tk.Button(win, text='Calcular Media', command=self.fCalcularMedia)
#            componentes TreeView        

        self.dadosColunas = ('Aluno', 'Nota1','Nota2', 'Media', 'Situacao')

        self.treeMedias = ttk.Treeview(win, columns=self.dadosColunas, selectmode='browse')
        self.verscrlbar = ttk.Scrollbar(win, orient="vertical", command=self.treeMedias.yview)
        self.verscrlbar.pack(side = 'right', fill = 'x')
        self.treeMedias.configure(yscrollcommand=self.verscrlbar.set)

        
#       posicionamento dos componentes na janela
        self.lblNome.place(x=100, y=50)
        self.txtNome.place(x=200, y=50)

        self.lblNota1.place(x=100, y=100)
        self.txtNota1.place(x=200, y=100)

        self.lblNota2.place(x=100, y=150)
        self.txtNota2.place(x=200, y=150)

        self.btnCalcular.place(x=100, y=200)

        self.treeMedias.place(x=100, y=300)
        self.verscrlbar.place(x=805, y=300, height=255)

        self.id = 0
        self.iid = 0

        self.carregarDadosIniciais()
#      carregar dados
    def carregarDadosIniciais (self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = pd.read_excel(fsave)
            print ('********** dados disponiveis **********')
            Print(dados)

            u=dados.count()
            print('u:'+str(u))
            nn=len(dados['Aluno'])
            for i in range (nn):
                nome = dados['Aluno'][i]
                nota1 = str(dados['Nota1'][i])
                nota2 = str(dados['Nota2'][i])
                media = str(dados['Media'][i])
                situacao = dados['Situacao'][i]

                self.treeMedias.insert('', 'end', iid=self.iid, values=(nome, nota1, nota2, media, situacao))
                self.iid = self.iid +1
                self.id = self.id +1
        except:
            print('Ainda não existem dados para carregar')
#       salvar dados
    def fSalvarDados(self):
        try:
            fsave = 'planilhaAlunos.xlsx'
            dados = []

            for line in self.treeMedias.get_children():
                lstDados = []
                
                for value in self.treeMedias.item(line) ['values']:
                    lstDados.append(value)

                dados.append(lstDados)

            df = pd.Dataframe(data=dados,columns=self.dadosColunas)

            planilha = pd.ExcelWriter(fsave)
            df.to_excel(planilha, 'dados', index=False)

            planilha.save()
            print('Dados Salvos')
        except:
            print('Nao foi possivel salvar os dados')
#    Calcular Media
    def fCalcularMedia(self):
        try:
            nome = self.txtNome.get()
            nota1 = float(self.txtNota1.get())            
            nota2 = float(self.txtNota2.get())
            media, situacao = self.fVerificarSituacao(nota1, nota2)
            self.treeMedias.insert('', 'end',iid=self.iid, values=(nome,str(nota1), str(nota2), str(media),situacao))
            self.iid = self.iid + 1
            self.id = self.id + 1
            self.fSalvarDados()
        except ValueError:
            print('Entre com valores validos')
        finally:
            self.txtNome.delete(0, 'end')
            self.txtNota1.delete(0, 'end')
            self.txtNota2.delete(0, 'end')

    def fVerificarSituacao(self, Nota1, Nota2):
        media=(Nota1+Nota2)/2
        if(media>=7.0):
            situacao = 'Aprovado'
        elif (media>=5.0):
            situacao = 'Em Recuperacao'
        else:
            situacao = 'Reprovado'
        return media, situacao
janela=tk.Tk()
principal=PrincipalRAD(janela)
janela.title('Bem Vindo ao RAD')
janela.geometry("820x600+10+10")
janela.mainloop()
