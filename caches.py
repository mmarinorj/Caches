import pandas as pd
import warnings as wn
import os
from PyQt5 import QtWidgets, uic
import sys


seg = ['D,R']
ter = ['T,AH']
qua = ['AJ,AX']
qui = ['AZ,BN']
sex = ['BP,CD']
sab = ['CF,CT']
dom = ['CV,DJ']
freelas = []
semana = []
totais = []
aba = 0
arquivo_fonte = ""
arquivo = ""


def salvar():
    global aba, freelas, semana, totais, seg, ter, qua, qui, sex, sab, dom, arquivo_fonte, arquivo
    while len(arquivo_fonte) < 1:
        arquivo_fonte = QtWidgets.QFileDialog.getOpenFileName(None, "Selecionar Fonte de Dados", "",
                                                              "XLSX files (*.xlsx)")[0]
        with open('aceobd' + '.cfg', 'w') as arquivo:
            arquivo.write(arquivo_fonte)
    if tela.radioButton.isChecked():
        aba = -2
    else:
        aba = -1
    semana.append(cache(seg, aba))
    semana.append(cache(ter, aba))
    semana.append(cache(qua, aba))
    semana.append(cache(qui, aba))
    semana.append(cache(sex, aba))
    semana.append(cache(sab, aba))
    semana.append(cache(dom, aba))

    for i in semana:
        for l in i:
            if len(l) > 0:
                freelas.append(l[0])

    freelas = list(set(freelas))
    freelas.sort()

    for f in freelas:
        tempx = 0
        for i in semana:
            for l in i:
                if len(l) > 0:
                    if f == l[0]:
                        tempx += l[1]
        totais.append([f, tempx])

    try:
        arquivo = QtWidgets.QFileDialog.getSaveFileName(None, "Salvar...", "",
                                                              "XLSX files (*.xlsx)")[0]
        dados = pd.DataFrame(data=totais)
        dados.to_excel(f'{arquivo}', index=False)
    except ValueError:
        pass
    except PermissionError:
        box.about(tela, 'Erro tentando salvar o arquivo!', f'O arquivo não pode ser salvo!\n'
                                                           f'Verifique se o mesmo está aberto para edição\n'
                                                           f'ou protegido contra gravação e tente novamente.')

    dialog = box.question(tela, "Sucesso!", f"O arquivo foi salvo com sucesso como:\n\n{arquivo}\n\nDeseja "
                                            f"Abrir o arquivo agora?",
                                            box.StandardButton.Yes | box.StandardButton.No)
    print(arquivo)
    if dialog != 65536:
        os.popen(arquivo)


def cache(dia, tab):
    global arquivo_fonte
    wn.simplefilter("ignore")
    df = pd.read_excel(arquivo_fonte, sheet_name=tab, usecols=dia[0])
    for j in range(df.shape[0]):
        temp = list(df.loc()[j])
        if pd.isnull(temp[0]):
            pass
        elif pd.isnull(temp[1]):
            pass
        elif temp[0] == 'EVENTO'\
                or temp[0] == 'MONTAGEM'\
                or temp[0] == 'ENTREGA MATERIAL'\
                or temp[0] == 'SEM USO'\
                or temp[0] == 'DESMONTAGEM':
            pass
        else:
            dia.append(temp)
    dia.pop(0)
    return dia


def sair_app():
    app.quit()


def fonte():
    global arquivo_fonte
    if os.path.isfile('aceobd.cfg'):
        with open('aceobd.cfg', 'r') as linha:
            arquivo_fonte = linha.readline().strip()
    if len(arquivo_fonte) > 0:
        dialog = box.question(tela, "Fonte de dados", f"O arquivo fonte de dados está selecionado"
                                                      f"no caminho:\n\n{arquivo_fonte}\n\n"
                                                      f"Deseja selecionar outro arquivo?",
                                                      box.StandardButton.Yes | box.StandardButton.No)
        if dialog != 65536:
            arquivo_fonte = QtWidgets.QFileDialog.getOpenFileName(None, "Selecionar Fonte de Dados", "",
                                                                  "XLSX files (*.xlsx)")[0]
            if len(arquivo_fonte) == 0:
                pass
            else:
                with open('aceobd' + '.cfg', 'w') as arquivo:
                    arquivo.write(arquivo_fonte)
    else:
        arquivo_fonte = QtWidgets.QFileDialog.getOpenFileName(None, "Selecionar Fonte de Dados", "",
                                                              "XLSX files (*.xlsx)")[0]
        if len(arquivo_fonte) == 0:
            pass
        else:
            with open('aceobd' + '.cfg', 'w') as arquivo:
                arquivo.write(arquivo_fonte)


def iniciar():
    global arquivo_fonte
    if os.path.isfile('aceobd.cfg'):
        with open('aceobd.cfg', 'r') as linha:
            arquivo_fonte = linha.readline().strip()
    if len(arquivo_fonte) > 0:
        pass
    else:
        box.about(tela, 'Arquivo não encontrado', 'O arquivo de dados ".XLSX" não foi encontrado.'
                                                  '\nSelecione o arquivo na próxima janela.\n')
        arquivo_fonte = QtWidgets.QFileDialog.getOpenFileName(None, "Selecionar Fonte de Dados", "",
                                                              "XLSX files (*.xlsx)")[0]
        with open('aceobd' + '.cfg', 'w') as arquivo:
            arquivo.write(arquivo_fonte)


### MAINWINDOW ###
app = QtWidgets.QApplication(sys.argv)
tela = uic.loadUi("main.ui")
tela.radioButton.setChecked(True)
tela.radioButton_2.setChecked(False)
tela.pushButton_2.clicked.connect(salvar)
tela.actionQuit.triggered.connect(sair_app)
tela.actionFonte_de_Dados.triggered.connect(fonte)
box = QtWidgets.QMessageBox

tela.show()
iniciar()
app.exec()
