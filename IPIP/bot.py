import time
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from openpyxl import load_workbook

class bot():
    def __init__(self):
        self.drive = webdriver.Chrome(executable_path=r"C:\Users\gleys\Desktop\IPIP\chromedriver.exe")
        self.wb = load_workbook(filename='dados.xlsx')
        self.ws = self.wb['Dados']

    def preencherForms(self):
        drive = self.drive
        drive.get('https://www.personal.psu.edu/~j5j/IPIP/ipipneo120.htm')
        time.sleep(3)
        yes1 = drive.find_element_by_xpath('/html/body/form/table[1]/tbody/tr/td/input')
        yes1.click()

        yes2 = drive.find_element_by_xpath('/html/body/form/table[2]/tbody/tr/td/input')
        yes2.click()

        botaoSend = drive.find_element_by_xpath('/html/body/form/b[2]/p/input')
        botaoSend.click()
        
        # esperar clicar no botao de aceitar sem seguranca
        time.sleep(20)

        # A pessoa que está sendo avaliada
        linhaDaPessoa = 2

        # nome
        nome = drive.find_element_by_xpath('/html/body/form/p[2]/input')
        nome.send_keys(self.ws[linhaDaPessoa][0].value)

        # sexo
        sexo = self.ws[linhaDaPessoa][2].value
        if (sexo == "Masculino"):
            sexoMasculino = drive.find_element_by_xpath('/html/body/form/p[3]/table/tbody/tr/td[2]/input')
            sexoMasculino.click()
        
        elif (sexo == "Feminino"):
            sexoFeminino = drive.find_element_by_xpath('/html/body/form/p[3]/table/tbody/tr/td[1]/input')
            sexoFeminino.click()

        # idade
        idade = drive.find_element_by_xpath('/html/body/form/p[4]/input')
        idade.send_keys(self.ws[linhaDaPessoa][3].value)

        # pais
        pais = drive.find_element_by_xpath('/html/body/form/p[6]/select')
        pais.send_keys('Brazil')

        
        for numeroQuestaoFormulario in range(1, 61):
            numeroQuestaoExcel = numeroQuestaoFormulario + 3
            respostaDaPessoa = self.ws[linhaDaPessoa][numeroQuestaoExcel].value
            qfinal = ""

            if (respostaDaPessoa == "Muito Impreciso"):
                qfinal = drive.find_element_by_xpath('/html/body/form/p[6]/table/tbody/tr['+str(numeroQuestaoFormulario)+']/td[3]/center/input')
                
            elif (respostaDaPessoa == "Moderadamente Impreciso"):
                qfinal = drive.find_element_by_xpath('/html/body/form/p[6]/table/tbody/tr['+str(numeroQuestaoFormulario)+']/td[4]/center/input')

            elif (respostaDaPessoa == "Nem preciso Nem impreciso"):
                qfinal = drive.find_element_by_xpath('/html/body/form/p[6]/table/tbody/tr['+str(numeroQuestaoFormulario)+']/td[5]/center/input')

            elif (respostaDaPessoa == "Moderadamente Preciso"):
                qfinal = drive.find_element_by_xpath('/html/body/form/p[6]/table/tbody/tr['+str(numeroQuestaoFormulario)+']/td[6]/center/input')

            elif (respostaDaPessoa == "Muito Preciso"):
                qfinal = drive.find_element_by_xpath('/html/body/form/p[6]/table/tbody/tr['+str(numeroQuestaoFormulario)+']/td[7]/center/input')      
            
            qfinal.click()

        botaoEnviarParte1 = drive.find_element_by_xpath('/html/body/form/p[7]/input')
        botaoEnviarParte1.click()        
        
        # esperar clicar no botao de aceitar sem seguranca
        time.sleep(60)



bot = bot()
bot.preencherForms()

