from selenium import webdriver
from selenium.webdriver.common.by import By
from time import sleep
from selenium.webdriver.support.ui import Select
import openpyxl

# Carregar planilha Excel
planilha_dados_consulta = openpyxl.load_workbook('dados da consulta.xlsx')
pagina_processos = planilha_dados_consulta['processos']

# Inicializando o WebDriver
driver = webdriver.Chrome()

# Acessando a página
driver.get('https://pje-consulta-publica.tjmg.jus.br/')

# Pausando por 2 segundos para garantir o carregamento inicial
sleep(2)

# Tentando localizar o elemento dentro da página principal
try:
    # Verifique se o elemento está dentro de um iframe
    iframes = driver.find_elements(By.TAG_NAME, "iframe")
    
    # Se houver iframes, mudar o foco para o correto
    if iframes:
        driver.switch_to.frame(iframes[0])  # Trocar para o primeiro iframe encontrado

    # Localizando o campo de input pelo ID
    elemento = driver.find_element(By.XPATH, "//input[@id='fPP:Decoration:numeroOAB']")

    # Rolando até o elemento estar visível
    driver.execute_script("arguments[0].scrollIntoView();", elemento)
    
    # Pausa para garantir que o scroll foi completado
    sleep(2)
    
    # Clicando no elemento
    elemento.click()
    sleep(1)

    # Digitando o número "259155"
    numero_oab = "259155"
    elemento.send_keys(numero_oab)
    sleep(1)

    # Selecionando o estado no dropdown
    selecao_uf = driver.find_element(By.XPATH, "//select[@id='fPP:Decoration:estadoComboOAB']")
    sleep(1)
    opcoes_uf = Select(selecao_uf)  
    sleep(1)
    opcoes_uf.select_by_visible_text('SP')  # Use select_by_visible_text para selecionar a UF
    sleep(1)

    # Localizando o botão de pesquisar e clicando
    botao_pesquisar = driver.find_element(By.XPATH, "//input[@id='fPP:searchProcessos']")
    botao_pesquisar.click()
    sleep(5)

    # Encontrando todos os links para abrir os detalhes dos processos
    links_abrir_processo = driver.find_elements(By.XPATH, "//a[@title='Ver Detalhes']")
    for link in links_abrir_processo:
        janela_principal = driver.current_window_handle
        link.click()
        sleep(5)
        
        # Trocando para a nova janela aberta
        janelas_abertas = driver.window_handles
        for janela in janelas_abertas:
            if janela != janela_principal:
                driver.switch_to.window(janela)
                sleep(5)
                
                # Coletando o número do processo
                numero_processo = driver.find_element(By.XPATH, "//div[@class='propertyView ']//div[@class='col-sm-12 ']").text
                
                # Coletando a lista de participantes
                participantes = driver.find_elements(By.XPATH, "//tbody[contains(@id,'processoPartesPoloAtivoResumidoList:tb')]//span[@class='text-bold']")
                lista_participantes = [participante.text for participante in participantes]

                # Verificando e salvando os dados na planilha
                if len(lista_participantes) == 1:
                    pagina_processos.append([numero_oab, numero_processo, lista_participantes[0]])
                else:
                    pagina_processos.append([numero_oab, numero_processo, ','.join(lista_participantes)])
                
                # Salvando a planilha
                planilha_dados_consulta.save('dados da consulta.xlsx')
                
                # Fechando a janela do processo
                driver.close()

        # Voltando para a janela principal
        driver.switch_to.window(janela_principal)

except Exception as e:
    print(f"Erro ao tentar rolar ou clicar no elemento: {e}")

# Fechando o driver
driver.quit()
