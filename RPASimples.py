import pyautogui as gui
from playsound import playsound
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
import sys
from win32api import GetKeyState
from win32con import VK_CAPITAL
import time
import unicodedata
import re
import unicodedata

#import logging
#import logging.config

"""[summary]

Returns:
    [type]: [description]
"""


class robo():
    _mensagem: str
    _falante: bool = True
    _titulo_dialogos: str = "SISTEMA RPA"
    _navegador: any
    _nome: str
    _log: any
    _dir_saida: str

    processo: str
    log: str
    pathvoz: str = './RPASimples/voz/'
    pathvoz_apresentacao: str
    path_webdriver: str
    pathlog: str = './LOG/'
    voz: any
    arquivo_log: any

    def __init__(self, nome: str, resolucao_x: int,
                 resolucao_y: int, pathvoz_apresentacao: str,
                 path_webdriver, dir_saida: str = "./",
                 nome_processo: str = ""):

        self.pathvoz_apresentacao = pathvoz_apresentacao
        self.path_webdriver = path_webdriver

        screenX, screenY = gui.size()
        if (screenX != resolucao_x or screenY != resolucao_y) and (resolucao_x != 0 and screenY != 0):
            gui.alert(
                f"Resolução do monitor não é compatível com este RPA ({resolucao_x} x {resolucao_y}) ")
            sys.exit()  # Fianliza o RPA

        gui.pause = 1
        capslock = GetKeyState(VK_CAPITAL)

        if capslock != 0:
            # Caso o capslock estiver ativo, ele será desligado
            gui.press('capslock')

        self._nome = nome
        self._dir_saida = dir_saida

        self.dir_arquivo_log = self._dir_saida + \
            f"//ARQUIVO_LOG_RPA_PROCESSO_{nome_processo}.txt"

        self.arquivo_log = open(self.dir_arquivo_log, "w")
        linha = f"Arquivo de logging processo {self._nome}\n"
        self.arquivo_log.write(linha)
        self.arquivo_log.close()

        self.set_mensagem("\n\nRPA ativado")
        assistencia = gui.confirm(text=f'Deseja manter a assistência por voz da {self._nome}?',
                                  title=f"{self._titulo_dialogos} - {nome}", buttons=['Sim', 'Não'])

        if assistencia == 'Sim':
            self._falante = True
        else:
            self._falante = False

        self.set_mensagem("Apresentação")
        self._interacao("APRESENTACAO")
        mensagem = """ATENÇÃO!!!  INICIANDO O RPA DE LANÇAMENTO {nome}
        VOCÊ TERÁ 10 SEGUNDOS PARA SE PREPARAR E USAR O COMPUTADOR ATÉ O FINAL DA EXECUÇÃO"""
        self.set_mensagem(mensagem)

        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')
        self.set_mensagem("Iniciando o processo")
        self.espera(10)

    def espera(self, segundos: int):
        self.set_mensagem("Espera "+str(segundos))
        time.sleep(segundos)

    def salvar_tela(self, nome_arquivo):
        self.set_mensagem("Print tela")
        tela = gui.screenshot(nome_arquivo)
        return tela

    def rolagem_tela(self, rodadas: int):
        self.set_mensagem("Rolagem de tela "+str(rodadas))
        gui.scroll(rodadas)

    def abrir_navegador(self, url: str):
        self.set_mensagem("Abrir navegador")
        chrome_options = Options()
        chrome_options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])
        chrome_options.add_argument('--lang=pr-BR')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_experimental_option(
            "excludeSwitches", ['enable-automation'])

        drive = f'{self.path_webdriver}chromedriver.exe'
        self._navegador = webdriver.Chrome(
            executable_path=drive,
            options=chrome_options)
        self._navegador.get(url)
        gui.press("F11")
        x, y = gui.size()
        gui.moveTo(x-15, 15)
        gui.click()
        self.espera(3)

    def fechar_navegador(self):
        gui.press("F11")
        # gui.hotkey('alt', 'f4')
        self._navegador.close()
        self._navegador.quit()
        self.set_mensagem("Navegador fechado")

    def abrir_link(self, url: str):
        self._navegador.get(url)
        self.set_mensagem(f"Link acessado: {url}")

    def escreva_gui(self, texto: str):
        self.set_mensagem("Escrevendo: "+texto)
        gui.typewrite(texto, interval=0.1)

    def click_elemento_web(self, xpath: str = "", nome_do_elemento: str = "") -> any:
        if xpath != "":
            elemento = self._navegador.find_element_by_xpath(xpath).click()
            self.set_mensagem(f"Elemento acessado xPath: {xpath}")
        elif nome_do_elemento != "":
            elemento = self._navegador.find_element_by_name(
                nome_do_elemento).click()
            self.set_mensagem(f"Elemento acessado xPath: {xpath}")
        else:
            self.set_mensagem(f"Elemento sem XPATH ou Nome")
        return elemento

    def combo_box_web(self, nome_do_elemento: str, texto_opcao_selecionada: str):
        elemento = self._navegador.find_element_by_xpath(
            f"//select[@name='{nome_do_elemento}']/option[text()='{texto_opcao_selecionada}']").click()
        return elemento

    def radio_box_web(self, nome_do_elemento: str, valor: str):
        elemento = self._navegador.find_element_by_xpath(
            ".//input[@type='{nome_do_elemento}' and @value='{valor}']").click()
        return elemento

    def entrar_dados_elemento_web(
            self, xpath: str, dados: str, key_enter: bool = False, key_tab: bool = False) -> any:

        elemento = self._navegador.find_element_by_xpath(xpath).click()
        elemento.send_keys(dados)

        if key_enter:
            elemento.send_keys(Keys.RETURN)
        if key_tab:
            elemento.send_keys(Keys.TAB)
        self.set_mensagem(
            f"Entrada de dados elemento: {xpath} | valor: {dados}")
        return elemento

    def set_mensagem(self, texto: str):
        self._mensagem = texto
        self._gravar_log()

    def set_titulo_dialogos(self, titulo):
        self._titulo_dialogos = titulo

    def _agora(self) -> str:
        data_e_hora_atuais = datetime.now()
        data_e_hora_em_texto = data_e_hora_atuais.strftime("%d/%m/%Y %H:%M")
        return data_e_hora_em_texto

    def _gravar_log(self):
        agora = self._agora()
        txt_log = f"{agora} | {self._mensagem}\n"

        self.arquivo_log = open(self.dir_arquivo_log, "a")
        self.arquivo_log.write(txt_log)
        self.arquivo_log.close()

    def bip(self, repetir):
        for i in range(repetir):
            playsound(self.pathvoz+'bip.mp3')

    def _interacao(self, envento: str):
        hora = datetime.now()
        hora = int(hora.strftime("%H"))

        if self._falante:
            if envento == "APRESENTACAO":
                if hora > 0 and hora < 12:
                    playsound(self.pathvoz+'bomdia.mp3')

                    if hora > 6 and hora < 9:
                        playsound(self.pathvoz +
                                  'esperoqueocafedamanhatenhasidobom.mp3')

                elif hora >= 12 and hora < 18:
                    playsound(self.pathvoz+'boatarde.mp3')

                    if hora >= 13 and hora < 14:
                        playsound(self.pathvoz+'esperoquejatenhaalmocado.mp3')

                else:
                    playsound(self.pathvoz+'boanoite.mp3')

                playsound(self.pathvoz_apresentacao+'apresentacao.mp3')

    def erro(self):
        self._interacao("ERRO")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def atencao(self):
        self._interacao("ATENCAO")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def mensagem(self):
        self._interacao("MENSAGEM")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def trabalho_concluido(self):
        self.set_mensagem("Mensagem de Alerta")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def dialogo(self, botao1: str, bota2: str) -> str:
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def tecla_gui(self, tecla: str):
        self.set_mensagem("Tecla pressionada "+tecla)
        gui.press(tecla)

    def processar_conjuto_de_campos_web(self, campos: list):

        for campo in campos:
            entrar_dados_elemento_web(
                campo['xpath'], campo['dados'], campo['key_enter'], campo['key_tab'])
            self.set_mensagem(
                f"Digitar campo: {campo['nome']}, XPath: {campo['xpath']}, Dados {campo['dados']}")


# funções auxiliares
def remover_acentos_caracteres_especiais(palavra):

    # Unicode normalize transforma um caracter em seu equivalente em latin.
    nfkd = unicodedata.normalize('NFKD', palavra)
    palavraSemAcento = u"".join(
        [c for c in nfkd if not unicodedata.combining(c)])

    # Usa expressão regular para retornar a palavra apenas com números, letras e espaço
    return str(re.sub('[^a-zA-Z0-9 \\\]', '', palavraSemAcento))
