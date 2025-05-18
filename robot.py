from selenium import webdriver  # Para automatizar o navegador
from selenium.webdriver.common.by import By  # Para localizar elementos na página
from webdriver_manager.chrome import ChromeDriverManager  # Para gerenciar o driver do Chrome
from selenium.webdriver.chrome.service import Service  # Para configurar o serviço do Chrome
import time  # Para pausas no script
from selenium.webdriver.support.ui import Select
import pandas as pd
#from IPython.display import display
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions
from customtkinter import *
from tkinter import messagebox
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox

def importar_planilha():
    global caminho_pacientes
    global dicionario_pacientes
    caminho_pacientes = filedialog.askopenfilename(title="Selecione um arquivo Excel",filetypes=[("Excel", "*.xlsx")])  
    if caminho_pacientes:
        try:
            ctk.set_appearance_mode("dark")
            caminho_pacientes_leitura = pd.read_excel(caminho_pacientes)
            CTkMessagebox( title="Sucesso!", message="Operação concluída com sucesso!", icon="check", option_1="OK")
            caminho_pacientes_formatacao = caminho_pacientes_leitura[['codigo', 'Especialidade', 'Profissional', 'Observação', 'Tipo Consulta']]
            dicionario_pacientes = caminho_pacientes_formatacao.to_dict()
        except Exception as e:
            messagebox.showerror("Erro", f" Não foi possível ler o arquivo:\n{e}")
importar_planilha

def iniciar_siresp():
    global navegador
    servico = Service(ChromeDriverManager().install()) # Instala automaticamente o ChromeDriver
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://www.siresp.saude.sp.gov.br")
    navegador.find_element('xpath', '//*[@id="btn-4"]').click()
    if navegador:
        CTkMessagebox( title="Sucesso!", message="Navegador aberto com sucesso!", icon="check", option_1="OK")
iniciar_siresp


                    
def Iniciar_robot(): 
    
    def aba_agendamento():
        lista_menu = navegador.find_elements('class name', 'sf-with-ul')
        for botao in lista_menu:
            if 'Agendamento' in botao.text:
                print("Agendamento Localizado")
                botao.click()
                
                # Localizando submenu de Agenda
                lista_agendamento = navegador.find_elements('xpath', '//a[contains(@href, "age_marcar.php?P=age_marcar")]')
                print("Agenda Localizada")
                
                for botao_agenda in lista_agendamento:
                    print(botao_agenda.text)
                    if 'Agenda' in botao_agenda.text:
                        print("Elemento Localizado")
                        botao_agenda.click()
                    else:
                        print("Elemento não encontrado")
    

      
    for indice in range(len(dicionario_pacientes)):
        time.sleep(3)  
        # Passo 5: Navegando no menu de Agendamento
        # Localizando todos os elementos do menu principal
        def retornando_menu():
            try:
                navegador.refresh()
                print('Atualizando a página')
                return True
            except:
                print('Erro ao atualizar a página')
                print('Passando para o próximo paciente')
                time.sleep(5)
                return False
        if not retornando_menu():
            print('Erro ao atualizar a página')
            print('Atualizando Pagina')
            print('Recomecando Busca')
            navegador.refresh()
            time.sleep(5) # Espera para carregamento
            retornando_menu()
        time.sleep(3)
        
        iframe = navegador.find_element('id', 'site')
        navegador.switch_to.frame(iframe)
        aba_agendamento()
        time.sleep(5) # Espera para carregamento

        def Buscar_Pac():
            try:    
                # Mudando para o iframe principal
                iframe2 = navegador.find_element('id', 'principal')
                navegador.switch_to.frame(iframe2)
                print("Iframe Localizado")
                codigo = int(dicionario_pacientes['codigo'][indice])
                print(codigo)
                # Inserindo código do paciente e buscando
                colocando_cod = navegador.find_element('id', 'FLT_COD_PACIENTE').send_keys(f"{codigo}")
                print('Código inserido')
                buscando = navegador.find_element('name', 'btn_acao').click()
                print('Botao selecionado')
                return True
            except:
                print("Erro ao buscar paciente")
                print("Passando para o próximo paciente")
                return False
        time.sleep(3)
        if not Buscar_Pac():
            print("Erro ao buscar paciente")
            print("Recomeçando busca")
            navegador.refresh()
            time.sleep(4) # Espera para carregamento
            retornando_menu()

        # Função para selecionar a aba de especialidades
        def aba_selecao():
            try:
                menu_selecao = navegador.find_element('id', 'c_filtro').click()
                print("Menu Selecionado")
                time.sleep(3)
                return True
            except:
                print("Menu não encontrado, ou conexão lenta")
                return False  
        if not aba_selecao():
            print("Erro ao selecionar menu")
            navegador.refresh()
            time.sleep(4)   
            retornando_menu()
        time.sleep(3)

        # Função para selecionar a especialidade médica
        """def sel_agenda():
            tabelasColunas = navegador.find_element('id', 'layer_espec').find_elements('id', 'ID_ESPECIALIDADE_ARR')
            divprinc = navegador.find_element(By.ID, 'layer_espec')
            
            for chk in tabelasColunas:
                valor = chk.get_attribute("value")
                try:
                    camId = divprinc.find_element(By.ID, valor) 
                    try:
                        print('valor do id encontrado', valor)
                        valor = camId.find_element(By.TAG_NAME, 'b')
                        navegador.execute_script("arguments[0].scrollIntoView();", valor)
                        print(valor.text)
                        
                        # Verificando se é a especialidade desejada
                        if valor.text.strip() == "CARDIOLOGIA":
                            print("Valor encontrado")
                            print("Especialidade encontrada")
                            chk.click()
                            break
                        else:
                            print("Valor não encontrado")
                    except:
                        print(f"Elemento Tag Name:b, com o ID:{valor} não encontrado")
                except:
                    print('Elemento ID não existe nesse momento!')

        sel_agenda()"""
        time.sleep(5)

        # Função para abrir a aba de seleção de médico
        def aba_medico():
            selecao = navegador.find_element(By.ID, 'c_profissional').click()
        aba_medico()
        time.sleep(6)

        # Função para selecionar o médico
        def sel_medico():
            try:
                tabelasColunas = navegador.find_element(By.ID, 'layer_prof').find_elements(By.ID, 'ID_PROFISSIONAL_ARR')
                especialidade = str(dicionario_pacientes['Especialidade'][indice]).upper()
                profissional = str(dicionario_pacientes['Profissional'][indice]).upper()
                print(profissional, especialidade)
                for chk in tabelasColunas:
                    valor = chk.get_attribute("value")
                    textovalor = f"{valor}"
                    valorFiltrado = textovalor.replace("|", "")
                    print(valorFiltrado)
                    camId = navegador.find_element(By.ID, valorFiltrado)
                    print('valor do id encontrado', valorFiltrado)
                    valorid = camId.find_element(By.XPATH, f'//td[@id="{valorFiltrado}"]')
                    navegador.execute_script("arguments[0].scrollIntoView();", valorid)
                    
                    # Limpando e formatando o texto
                    texto_medico_especialidade = (f"{especialidade} {profissional}")
                    textostringvalor = valorid.text
                    texto_limpo = textostringvalor.strip().replace("&nbsp;", " ").replace("\n", " ").replace("\t", " ")
                    print(texto_medico_especialidade)
                    print(texto_limpo)
                    # Verificando médico e especialidade
                    if texto_medico_especialidade == texto_limpo: #and f"{profissional}" in texto_limpo:
                        print("Valor encontrado")
                        print("Medico encontrado")
                        chk.click()
                        return True
                    else:
                        print("Valor não encontrado")
                print("Passando para o próximo paciente")
                return False
            except:
                print('Especialidade não disponível')
                print('Passando para o proximo paciente')
                return False
        time.sleep(2)
        if not sel_medico():
            print('Erro ao selecionar médico')
            print('Passando para o próximo paciente')
            continue
        time.sleep(3)
        # Função para selecionar agenda interna (comentada)
        def aba_agInt():
            try:
                selecaoint = navegador.find_element('id', 'c_agenda_int').click()
            except:
                print("Erro ao selecionar agenda interna")
                print("Passando para o próximo paciente")
        time.sleep(3)
        aba_agInt()
        
        def tipo_consulta():
            try:    
                iframedia = navegador.find_element(By.ID, 'agenda_ce')
                navegador.switch_to.frame(iframedia)
                time.sleep(4)
                tip_consulta = str(dicionario_pacientes['Tipo Consulta'][indice]).upper()
                print(tip_consulta)
                #input('Pressione Enter para continuar...')
                opcinter = 'INTERCONSULTA'
                #opcretorno = 'Retorno'
                # Selecionando o tipo de consulta
                if tip_consulta == opcinter:
                    time.sleep(4)
                    navegador.find_element(By.NAME, 'FLT_INTERCONSULTA_CHK').click()
                    print("Tipo de consulta selecionado : INTERCONSULTA")
                    return
                else:
                    time.sleep(4)
                    navegador.find_element(By.NAME, 'FLT_RETORNO_CHK').click()
                    print("Tipo de consulta selecionado : RETORNO")
                    return
                #input("Pressione Enter para continuar...") #esperando o usuario preencher a captcha
            except:
                print("Não Ha mais cotas para esta especialidade")
                print('passando para o proximo paciente')
        time.sleep(5)
        tipo_consulta()
        
        # Seção final comentada para seleção de dia na agenda
        def sel_diaAg():
            print("Selecionando dia")
            #iframedia = navegador.find_element(By.ID, 'agenda_ce')
            #navegador.switch_to.frame(iframedia)
            print("Iframe Agenda Localizado")
            #selecionando Tipo de Consulta
            #selConsulta = navegador.find_element(By.XPATH, '//tr[@class="dados"]//td')
            #divprinc = navegador.find_element(By.CLASS_NAME, 'div') 
            div_mes0 = navegador.find_elements(By.XPATH, '//div[@id="div_mes0"]//td')
            div_mes1 = navegador.find_elements(By.XPATH, '//div[@id="div_mes1"]//td')
            div_mes2 = navegador.find_elements(By.XPATH, '//div[@id="div_mes2"]//td')
            div_mes3 = navegador.find_elements(By.XPATH, '//div[@id="div_mes3"]//td')
            div_mes4 = navegador.find_elements(By.XPATH, '//div[@id="div_mes4"]//td')
            
            try:
                try:
                    for percorre in div_mes0:
                        #time.sleep(0.5)
                        mes = navegador.find_element(By.ID, 'mes0')
                        print(f"Percorrendo agenda de {mes.text}")
                        verificando = percorre.get_attribute("class")
                        print(verificando)
                        # Verifica se o texto contém a data desejada
                        condicao = "dia dados p" == verificando
                        condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                        if condicao or condicao2 == verificando:
                            print("Dia encontrado")
                            # Clica no dia desejado
                            clicou = percorre.click()
                            return True
                        else :
                            print("procurando próximo dia disponivel")
                    navegador.find_element(By.ID, 'mes1').click()   
                except:
                    print('passando para o proximo paciente') 
                    return False
                
                # Percorrendo os dias do mês 1
                try:
                    for percorre in div_mes1:
                        #time.sleep(0.5)
                        mes = navegador.find_element(By.ID, 'mes1')
                        print(f"Percorrendo agenda de {mes.text}")
                        verificando = percorre.get_attribute("class")
                        print(verificando)
                        condicao = "dia dados p" == verificando
                        condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                        # Verifica se o texto contém a data desejada
                        if condicao or condicao2:
                            print("Dia encontrado")
                            # Clica no dia desejado
                            percorre.click()
                            return True
                        else :
                            print("Dia não encontrado, procurando próximo dia")
                            
                    navegador.find_element(By.ID, 'mes2').click()
                    
                except:
                    print('passando para o proximo paciente')
                    return False 
                try:
                    for percorre in div_mes2:
                    # time.sleep(0.5)
                        mes = navegador.find_element(By.ID, 'mes2')
                        print(f"Percorrendo agenda de {mes.text}")
                        verificando = percorre.get_attribute("class")
                        print(verificando)
                        condicao = "dia dados p" == verificando
                        condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                        # Verifica se o texto contém a data desejada
                        if condicao or condicao2:
                            print("Dia encontrado")
                            # Clica no dia desejado
                            percorre.click()
                            return True
                        else :
                            print("Dia não encontrado, procurando próximo")
                    
                    navegador.find_element(By.ID, 'mes3').click()
                except:
                    print('passando para o proximo paciente')
                    return False   
                
                try:            
                    for percorre in div_mes3:
                        mes = navegador.find_element(By.ID, 'mes3')
                        print(f"Percorrendo agenda de {mes.text}")
                        verificando = percorre.get_attribute("class")
                        print(verificando)
                        condicao = "dia dados p" == verificando
                        condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                        # Verifica se o texto contém a data desejada
                        if condicao or condicao2:
                            print("Dia encontrado")
                            # Clica no dia desejado
                            percorre.click()
                            return True
                            break
                            
                        else :
                            print("Dia não encontrado, procurando próximo")
                    
                    navegador.find_element(By.ID, 'mes4').click()
                except:
                    print('passando para o proximo paciente') 
                    return False
        
                
                try:       
                    for percorre in div_mes4:
                        mes = navegador.find_element(By.ID, 'mes4')
                        print(f"Percorrendo agenda de {mes.text}")
                        verificando = percorre.get_attribute("class")
                        print(verificando)
                        condicao = "dia dados p" == verificando
                        condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                        # Verifica se o texto contém a data desejada
                        if condicao or condicao2:
                            print("Dia encontrado")
                            # Clica no dia desejado
                            percorre.click()
                            return True
                            break
                        else :
                            print("Dia não encontrado, procurando próximo")
                    
                    print(f"Agendamento no mes: {mes.text} não disponível")
                except:
                    print('passando para o proximo paciente')
                    return False        
                            
            except:
                print("Agendamento Não Disponivel para esta especialidade")  
                print("Passando para o proximo paciente")
                return False
        if not sel_diaAg():
            print("Agendamento não disponível")
            print("Passando para o proximo paciente")
            continue
        time.sleep(3)     
                
                
        def selecionar_horario():
            time.sleep(5)
            print("Procurando Horario disponível")
            #clicar para abrir a caixa de opçoes
            horario_frame = navegador.find_element(By.ID, 'horario_frame')
            navegador.switch_to.frame(horario_frame)
            navegador.find_element(By.XPATH, "//select[contains(@name, 'CB_CONSULTA')]").click()
            caminho = navegador.find_element(By.XPATH, "//select[contains(@name, 'CB_CONSULTA')]")
            time.sleep(2)

            #drop.find_element(By.NAME,'CB_CONSULTA').click()
            #divsec = divprin.find_elements(By.TAG_NAME, 'td')
            #(By.XPATH, f'//td[@id="{valorFiltrado}"]')
            #clicando = navegador.find_element(By.NAME, 'CB_CONSULTA').click()

            dropdown = caminho.find_element(By.XPATH, "//select[contains(@name, 'CB_CONSULTA')]")
            roll = Select(dropdown)
            listaopc = roll.options[0:] # Pega todas as opções
            primopc= listaopc[1]
            #navegador.execute_script("arguments[0].click();", primopc)
            #dropdown.select_by_index(1)
            if primopc:
                print("Espere 1 segundo, selecionando horario disponivel")
                time.sleep(1)
                primopc.click()
                print(primopc.text)
            else:
                print("Não ha horarios Disponiveis") 
        selecionar_horario()
        time.sleep(2)

        def comentando():
            try:    
                observacao = dicionario_pacientes['Observação'][indice]
                procurando_obs = navegador.find_element(By.NAME, 'AGE_OBSERVACAO')
                time.sleep(1)
                navegador.execute_script("arguments[0].scrollIntoView();", procurando_obs)
                procurando_obs.click()
                procurando_obs.send_keys(f"{observacao}//ROBOT")
            except:
                print("Não é possivel comentar")
                return
        comentando()
        #input("Pressione Enter para continuar...") #esperando o usuario preencher a captcha
        time.sleep(2)

        def botao_agendar_paciente():
            try:
                procurando_botao = navegador.find_element(By.NAME, 'marcar_c')
                procurando_botao.click()
            except:
                print("Erro ao Agendar o Paciente")
                print("Passando para o próximo paciente")
                saindo_da_impress()   
        botao_agendar_paciente()
        time.sleep(1)
        def alert_a():
            try:
                print("Procurando Alerta")   
                time.sleep(2)
                print("Clicando no alerta")
                div_widget = navegador.find_element(By.XPATH, 'u/html/body/div')
                print('Div Widget encontrado')
                time.sleep(2)
                div_protocolo = div_widget.find_element(By.XPATH, '//*[@id="divMostraProtocolo"]')
                print('Div Protocolo encontrado')
                time.sleep(2)
                div_protocolo.find_element(By.XPATH, '//*[@id="btn_incluir"]').click()
                print("Alerta encontrado")
            except:
                return
        alert_a()
        
        time.sleep(3)
        def saindo_da_impress():
            try:
                alerta = navegador.switch_to.alert
                alerta.dismiss()  # Clica em "Cancelar"
                print("Alerta cancelado")
                print("Retornando ao agendamento")
            except:
                return
        saindo_da_impress()
        time.sleep(2)
input()


