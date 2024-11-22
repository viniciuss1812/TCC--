
from time import sleep
from DrissionPage import ChromiumPage
import pandas as pd
import pyautogui
import pyperclip
import PyPDF2
import os
import glob




csvarqv = pd.read_csv("ARQUIVOPESSOA.csv", sep=";", encoding="Windows-1252")
csvarqv.columns = csvarqv.columns.str.strip()  # Remove espaços extras dos nomes das colunas


pd.set_option('display.max_columns', None)
pd.set_option('display.expand_frame_repr', False)


print("Conteúdo do DataFrame:")
print(csvarqv.to_string(index=False))

print("\nNomes das Colunas:", csvarqv.columns.tolist())




try:
    for index, row in csvarqv.iterrows():
        CPF = row['CPF']
        Nome = row['Nome']
        Nacionalidade = row['Nacionalidade']
        DataNascimento = row['Data de Nascimento']
        PaisNascimento = row['País de nascimento']
        UfNascimento = row['UF de nascimento']
        Municipio = row['Município de nascimento']
        Municipioteste = 'Estância';
        Pai = row['Nome do pai']
        Mae = row['Nome da mãe']
        Tipo = row['Tipo']
        N = row['N° *']
        OrgaoExpedicao =row['Órgão de expedição *']
        UfExpedicao = row['UF de expedição *']

        def extrair_duas_palavras(frase, palavras_procuradas):
    
           if palavras_procuradas in frase:
              return f'{Nome} não apresenta Antecedentes Criminais'
           else:
              return f'{Nome} apresenta Antecedentes Crimanais'

        # Enviar os valores para a aplicação ativa
        print(f"Enviando dados da linha {index + 1}: Nome = {Nome}, CPF = {CPF}")
        p = ChromiumPage()
        p.get('https://servicos.pf.gov.br/epol-sinic-publico/')
        sleep(25)
        xcapthca = 423; ycaptcha = 470;
        xnacionalidade = 1383; ynacionalidade = 657;
        xconfirmanacionalidade = 1382; yconfirmanacionalidade = 799;
        xsai = 1718; ysai = 771;
        xnascimento = 707; ynascimento = 768;
        x = 728; y = 655
        xuf = 1377; yuf= 765;
        xufconfima = 1427; yufconfirma = 799;
        xmunicipio  = 716; ymunicipio  = 865;
        xmunicipioconfima  = 738; ymunicipioconfirma  = 897;
        xpais = 1091; ypais = 792;
        # posicao = pyautogui.position()
        # print(f"Posição capturada: x={posicao.x}, y={posicao.y}")
        pyautogui.click(xcapthca, ycaptcha)

        sleep(15)
        pyautogui.click(x, y)
        sleep(8)

        pyautogui.write(str(CPF))
   
        pyautogui.press('tab')
        pyautogui.write(str(Nome))
        sleep(10)
        pyautogui.click(xnacionalidade, ynacionalidade)
        sleep(5)
        pyautogui.write(str(Nacionalidade))
        sleep(5)
        pyautogui.click(xconfirmanacionalidade, yconfirmanacionalidade)
        sleep(5)
        pyautogui.click(xsai, ysai)
        sleep(5)
        pyautogui.click(xnascimento, ynascimento)
        pyautogui.write(str(DataNascimento))
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(str(PaisNascimento))
        sleep(8)
        pyautogui.click(xpais, ypais)

        pyautogui.click(xuf, yuf)
        sleep(8)

        pyautogui.write(str(UfNascimento))
        sleep(10)
        pyautogui.click(xufconfima, yufconfirma)
        sleep(5)
        pyautogui.click(xmunicipio, ymunicipio)
        sleep(10)

        pyperclip.copy(Municipioteste)
        pyautogui.hotkey('ctrl', 'v')
        #pyautogui.write(str(Municipioteste))
        sleep(10)
        pyautogui.click(xmunicipioconfima, ymunicipioconfirma)
        sleep(10)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(str(Pai))
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(str(Mae))
        sleep(5)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.write(str(Tipo))
        sleep(10)
        pyautogui.press('down')
        pyautogui.press('enter')
        pyautogui.press('tab')
        pyautogui.press('tab')
        sleep(5)
        pyautogui.write(str(N))
        pyautogui.press('tab')
        pyautogui.write(str(OrgaoExpedicao))
        pyautogui.press('tab')
        sleep(5)
        pyautogui.write(str(UfExpedicao))
        sleep(4)
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('tab')
        pyautogui.press('enter')

        sleep(10)
        p.close()
        #fechar o botão e fechar o navegador
        lista_arquivos = glob.glob(r"C:\Users\vinic\Downloads\*") 
        arquivoatual = max(lista_arquivos, key=os.path.getmtime)
        caminhocompleto = os.path.join(os.path.dirname(arquivoatual), os.path.basename(arquivoatual))
        print(caminhocompleto)

        with open(caminhocompleto, "rb") as arquivo:
            leitorpdf = PyPDF2.PdfReader(arquivo)
            texto = ""
            for pagina in leitorpdf.pages:
              texto += pagina.extract_text()

        print("Conteúdo do PDF extraído:")
        print(texto)

        frase = texto
        palavras = "NÃO"  # Procurando pela sequência exata "não CONSTA"
        resultado = extrair_duas_palavras(frase, palavras)
        print(resultado)

        nome_arquivo = "Resultados.txt"

        with open(nome_arquivo, "a", encoding="utf-8") as arquivo:
            arquivo.write(resultado)

        print(f"O arquivo '{nome_arquivo}' foi preenchido com sucesso.")



except KeyError as e:
    print(f"\nExceção: A coluna {e} não foi encontrada no DataFrame.")
except Exception as ex:
    print(f"\nErro inesperado: {ex}")
        



 



# p.get('https://servicos.pf.gov.br/epol-sinic-publico/')

#  # Acessa o iframe desejado
# # i = p.get_frame('@src^https://challenges.cloudflare.com/cdn-cgi')
# # i = p.get_frame('cb-lb')
# # e = i('input[type="checkbox"]')
# # sleep(15)
# # e.click()
# # print ("PASSOU")
# sleep(50)

# pyautogui.press('tab')
# pyautogui.press('tab')
# pyautogui.press('tab') 
# pyautogui.press('tab')
# pyautogui.press('tab')

# pyautogui.write(CPF)
# pyautogui.press('tab')
# pyautogui.write(Nome)
# pyautogui.press('tab')







