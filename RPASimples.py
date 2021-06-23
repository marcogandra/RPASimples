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
import pandas as pd
import subprocess

# import logging
# import logging.config

"""[summary]

Returns:
    [type]: [description]
"""


class robo():
    __mensagem: str
    __titulo_dialogos: str = "SISTEMA RPA"
    __navegador: any
    __nome: str
    __log: any
    __dir_saida: str
    __winium: any
    __app: any

    processo: str
    log: str
    driver: str
    pathlog: str = './LOG/'
    arquivo_log: any
    arquivo_log_erros: any
    dir_arquivo_log: str
    dir_arquivo_log_erros: str

    def __init__(self,
                 nome: str,
                 resolucao_x: int = 0,
                 resolucao_y: int = 0,
                 driver: str = "./Driver/",
                 dir_saida: str = "./",
                 nome_processo: str = "",
                 pasta_de_downloads: str = "/Downloads/"):

        gui.FAILSAFE = False
        self.driver = driver
        self.pasta_de_downloads = pasta_de_downloads

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

        self.__nome = nome
        self.__dir_saida = dir_saida

        self.dir_arquivo_log = self.__dir_saida + \
            f"//ARQUIVO_LOG_RPA_PROCESSO_{nome_processo}.txt"

        self.dir_arquivo_log_erros = self.__dir_saida + \
            f"//ARQUIVO_LOG_RPA_PROCESSO_{nome_processo}_ERROS.txt"

        self.arquivo_log = open(self.dir_arquivo_log,
                                mode="w", encoding="utf-8")
        linha = f"Arquivo de logging processo {self.__nome}\n"
        self.arquivo_log.write(linha)
        self.arquivo_log.close()

        self.arquivo_log_erros = open(
            self.dir_arquivo_log_erros, mode="w", encoding="utf-8")
        linha = f"Arquivo de falhas processo {self.__nome}\n"
        self.arquivo_log_erros.write(linha)
        self.arquivo_log_erros.close()

        self.set_mensagem("\n\nRPA ativado")

        self.set_mensagem("Apresentação")
        mensagem = """ATENÇÃO!!!  INICIANDO O RPA EM 10 SEGUNDOS | FAVOR NÃO USAR/MINIMIZAR/BLOQUEAR O COMPUTADOR OU TERMINAL ATÉ O FINAL DA EXECUÇÃO"""
        self.set_mensagem(mensagem)

        gui.alert(text=self.__mensagem,
                  title=self.__titulo_dialogos, button='OK')
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

    def abrir_app(self, app_dir: str):
        self.__winium = subprocess.Popen(self.driver)

        self.__app = webdriver.Remote(
            command_executor='http://localhost:9999',
            desired_capabilities={
                'app': app_dir
            })

    def abrir_navegador(self, url: str):
        self.set_mensagem("Abrir navegador")
        chrome_options = Options()
        chrome_options.add_experimental_option(
            'excludeSwitches', ['enable-logging'])
        chrome_options.add_argument('--lang=pr-BR')
        chrome_options.add_argument('--disable-notifications')
        chrome_options.add_experimental_option(
            "excludeSwitches", ['enable-automation'])

        chrome_options.add_experimental_option('prefs', {
            "plugins.plugins_list": [{"enabled": False, "name": "Chrome PDF Viewer"}],
            "download": {
                "prompt_for_download": False,
                "default_directory": self.pasta_de_downloads
            }
        })

        drive = f'{self.path_webdriver}chromedriver.exe'

        self.__navegador = webdriver.Chrome(
            executable_path=drive,
            options=chrome_options)
        self.__navegador.get(url)
        gui.press("F11")
        x, y = gui.size()
        gui.moveTo(x-15, 15)
        gui.click()

    def fechar_app(self):
        if self.__app:
            self.__app.close()
            self.__app.quit()
        # gui.hotkey("ALT", "F4")   
        self.set_mensagem("Aplicativo fechado")

    def fechar_navegador(self):
        gui.press("F11")

        if self.__navegador:
            self.__navegador.close()
            self.__navegador.quit()

        self.set_mensagem("Navegador fechado")

    def abrir_link(self, url: str):
        self.__navegador.get(url)
        self.set_mensagem(f"Link acessado: {url}")

    def escreva_gui(self, texto: str):
        self.set_mensagem("Escrevendo: "+texto)
        gui.typewrite(texto, interval=0.1)

    def click_elemento_web(self, xpath: str = "", nome_do_elemento: str = "") -> any:
        if xpath != "":
            elemento = self.__navegador.find_element_by_xpath(xpath).click()
            self.set_mensagem(f"Elemento acessado xPath: {xpath}")
        elif nome_do_elemento != "":
            elemento = self.__navegador.find_element_by_name(
                nome_do_elemento).click()
            self.set_mensagem(f"Elemento acessado xPath: {xpath}")
        else:
            self.set_mensagem(f"Elemento sem XPATH ou Nome")
        return elemento

    def combo_box_web(self, nome_do_elemento: str, texto_opcao_selecionada: str):
        elemento = self.__navegador.find_element_by_xpath(
            f"//select[@name='{nome_do_elemento}']/option[text()='{texto_opcao_selecionada}']").click()
        return elemento

    def radio_box_web(self, nome_do_elemento: str, valor: str):
        elemento = self.__navegador.find_element_by_xpath(
            ".//input[@type='{nome_do_elemento}' and @value='{valor}']").click()
        return elemento

    def entrar_dados_elemento_web(
            self, xpath: str, dados: str, key_enter: bool = False, key_tab: bool = False) -> any:

        elemento = self.__navegador.find_element_by_xpath(xpath).click()
        elemento.send_keys(dados)

        if key_enter:
            elemento.send_keys(Keys.RETURN)
        if key_tab:
            elemento.send_keys(Keys.TAB)
        self.set_mensagem(
            f"Entrada de dados elemento: {xpath} | valor: {dados}")
        return elemento

    def set_mensagem(self, texto: str):
        self.__mensagem = texto
        self.__gravar_log()

    def set_titulo_dialogos(self, titulo):
        self.__titulo_dialogos = titulo

    def __agora(self) -> str:
        data_e_hora_atuais = datetime.now()
        data_e_hora_em_texto = data_e_hora_atuais.strftime("%d/%m/%Y %H:%M")
        return data_e_hora_em_texto

    def __gravar_log(self):
        agora = self.__agora()
        txt_log = f"{agora} | {self.__mensagem}\n"

        self.arquivo_log = open(self.dir_arquivo_log,
                                mode="a", encoding="utf-8")
        self.arquivo_log.write(txt_log)
        self.arquivo_log.close()

    def erro(self):
        self.arquivo_log_erros = open(
            self.dir_arquivo_log_erros, mode="a", encoding="utf-8")
        linha = f"Arquivo de falhas processo {self.__nome}\n"
        self.arquivo_log_erros.write(self.__mensagem)
        self.arquivo_log_erros.close()

    def atencao(self):
        gui.alert(text=self.__mensagem,
                  title=self.__titulo_dialogos, button='OK')

    def mensagem(self):
        gui.alert(text=self.__mensagem,
                  title=self.__titulo_dialogos, button='OK')

    def trabalho_concluido(self):
        self.set_mensagem("Mensagem de Alerta")
        gui.alert(text=self.__mensagem,
                  title=self.__titulo_dialogos, button='OK')

    def dialogo(self, botao1: str, bota2: str) -> str:
        gui.alert(text=self.__mensagem,
                  title=self.__titulo_dialogos, button='OK')

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
    nfkd = unicodedata.normalize('NFKD', str(palavra))
    palavraSemAcento = u"".join(
        [c for c in nfkd if not unicodedata.combining(c)])

    # Usa expressão regular para retornar a palavra apenas com números, letras e espaço
    return str(re.sub('[^a-zA-Z0-9 \\\]', '', palavraSemAcento))


def list_dict_para_excel(list_dict: dict, arquivo: str):
    df = pd.DataFrame(list_dict)
    writer = pd.ExcelWriter(arquivo, engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    writer.save()
    
def encoding_arquivo(arquivo):
    with open(arquivo, 'rb') as file:
    content = file.read()

    suggestion = UnicodeDammit(content)    
    
    return suggestion.original_encoding
