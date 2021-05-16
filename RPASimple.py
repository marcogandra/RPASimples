import pyautogui as gui
from playsound import playsound
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

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

    processo: str
    log: str
    pathvoz: str = './voz/'
    pathlog: str
    voz: any

    def __init__(self, nome):
        agora = self._agora()
        self._nome = nome
        self.pathlog = f'./{agora}_log/'
        self.set_mensagem("RPA ativado")
        assistencia = gui.confirm(text=f'Deseja manter a assistência por voz da {self._nome}?',
                                  title=f"{self._titulo_dialogos} - {nome}", buttons=['Sim', 'Não'])

        if assistencia == 'Sim':
            self._falante = True
        else:
            self._falante = False

    def abrir_navegador(self, url: str):
        self._navegador = webdriver.Chrome()
        self._navegador.get(url)
        gui.press("F11")
        x, y = gui.size()
        gui.moveTo(x-15, 15)
        gui.click()

    def fechar_navegador(self):
        gui.press("F11")
        self._navegador.close()

    def abrir_link(self, url: str):
        self._navegador.get(url)

    def click_elemento_web(self, xpath: str) -> any:
        elemento = self._navegador.find_element_by_xpath(xpath).click()
        return elemento

    def entrar_dados_elemento_web(
            self, xpath: str, dados: str, key_enter: bool = False, key_tab: bool = False) -> any:

        elemento = self._navegador.find_element_by_xpath(xpath).click()
        elemento.send_keys(dados)

        if key_enter:
            elemento.send_keys(Keys.RETURN)
        if key_tab:
            elemento.send_keys(Keys.TAB)

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
        txt_log = f"{agora} | {self._mensagem}"
        print(txt_log)

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

                playsound(self.pathvoz+'apresentacao.mp3')

    def apresentacao(self):
        self.set_mensagem("Apresentação")
        self._interacao("APRESENTACAO")
        self.set_mensagem("Mensagem de Alerta")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')
        self.set_mensagem("Iniciando o processo")

    def erro(self):
        self._interacao("ERRO")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def atencao(self):
        self._interacao("ATENCAO")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def log(self):
        pass

    def trabalho_concluido(self):
        self.set_mensagem("Mensagem de Alerta")
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def dialogo(self, botao1: str, bota2: str) -> str:
        gui.alert(text=self._mensagem,
                  title=self._titulo_dialogos, button='OK')

    def tecla_gui(self, tecla):
        gui.press(tecla)


luzbot = robo("Luzbot")
#luzbot.set_titulo_dialogos("RPA DARF")
# luzbot.apresentacao()
luzbot.abrir_navegador(
    "https://depositojudicial.caixa.gov.br/sigsj_internet/principal.xhtml")
gui.alert("x")
luzbot.click_elemento_web('//*[@id="tipoJusticaDeposito"]')
luzbot.tecla_gui("DOWN")
luzbot.tecla_gui("DOWN")
luzbot.tecla_gui("ENTER")
luzbot.click_elemento_web(
    '//*[@id="btnConfirmarDepositosJudiciais"]')
luzbot.click_elemento_web(
    '//*[@id="j_id45:filtroView:j_id47:tpDeposito1"]/tbody/tr/td/span/span/img')
luzbot.click_elemento_web(
    '//*[@id="j_id45:filtroView:j_id47:btnConfirmar"]/img')
luzbot.fechar_navegador()
