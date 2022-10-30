from ast import Break, While
from botcity.core import DesktopBot
import pandas as pd 

planilha = pd.read_excel('Transferencia.xlsx')


class Bot(DesktopBot):
    def action(self, execution=None):
     

        #Abrir o Protheus 
        self.browse(r"C:\Users\botcity.admin\Desktop\Protheus.lnk")
        #cofido para dar erro na Bot city  - SUCESS - 
        if not self.find( "OK", matching=0.97, waiting_time=90000):
            self.not_found("OK")
        self.click()     

        #usuraio e senha 
        if not self.find( "Usuario", matching=0.97, waiting_time=90000):
            self.not_found("Usuario")
        self.click_relative(63, 35)
        ##usuario
        self.paste('')
        self.tab()
        ##senha
        self.paste('')
        if not self.find( "Entrar", matching=0.97, waiting_time=10000):
            self.not_found("Entrar")
        self.click()       
        #preparar o ambiente     
        if not self.find( "bemvindo", matching=0.97, waiting_time=10000):
            self.not_found("bemvindo")
        self.wait(15000)
        self.tab()
        self.tab()
        self.paste('00601000')
        ##modulo financeiro
        self.paste('06')
        if not self.find( "Entrar", matching=0.97, waiting_time=10000):
            self.not_found("Entrar")
        self.click()
        
        #entrar na rotina 
        if not self.find( "Favoritos", matching=0.97, waiting_time=90000):
            self.not_found("Favoritos")
        self.click()
        if not self.find( "funções", matching=0.97, waiting_time=10000):
            self.not_found("funções")
        self.click()
        if not self.find( "confirmar", matching=0.97, waiting_time=10000):
            self.not_found("confirmar")
        self.click()
        
        #filtro
        if not self.find( "filtro", matching=0.97, waiting_time=900000):
            self.not_found("filtro")
        self.click()
        if not self.find( "coluna", matching=0.97, waiting_time=90000):
            self.not_found("coluna")
        self.click()
        if not self.find( "click_rolar", matching=0.97, waiting_time=90000):
            self.not_found("click_rolar")
        self.triple_click()
        self.wait(3000)
        if not self.find( "cliente", matching=0.97, waiting_time=10000):
            self.not_found("cliente")
        self.click()
        self.wait(3000)
        self.space()
        for codigo in planilha['Cliente']:
            self.wait(20000)
            if not self.find( "procurar", matching=0.97, waiting_time=10000):
                self.not_found("procurar")
            self.click_relative(84, 11)
            self.paste(f'{codigo}'.zfill(6))
            if not self.find( "Lupa", matching=0.97, waiting_time=10000):
                self.not_found("Lupa")
            self.click()
            self.wait(20000)
            while True: 
                if not self.find( "trans", matching=0.97, waiting_time=10000):
                    self.not_found("trans")
                self.click()
                while not self.find( "transferir1", matching=0.97, waiting_time=10000):
                    if not self.find( "trans", matching=0.97, waiting_time=10000):
                        self.not_found("trans")
                    self.click()
                if not self.find( "transferir1", matching=0.97, waiting_time=10000):
                    self.not_found("transferir1")
                self.click()
                if self.find( "Fechar", matching=0.97, waiting_time=10000):
                    self.click()
                    break
                else:
                    if not self.find( "portador", matching=0.97, waiting_time=10000):
                        self.not_found("portador")
                    self.click_relative(-51, 10)
                    self.paste('033')
                    self.wait(5000)
                    self.paste('4620')
                    self.tab()
                    self.paste('130046637')
                    self.wait(3000)
                    self.tab()
                    self.paste('4')
                    self.wait(8000)
                    if not self.find( "OKK", matching=0.97, waiting_time=20000):
                        self.not_found("OKK")
                    self.click()
                    while not self.find( "sim", matching=0.97, waiting_time=20000):
                        if not self.find( "OKK", matching=0.97, waiting_time=20000):
                            self.not_found("OKK")
                        self.click()
                    if not self.find( "sim", matching=0.97, waiting_time=20000):
                        self.not_found("sim")
                    self.click()
                    self.wait(25000)

        

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()


