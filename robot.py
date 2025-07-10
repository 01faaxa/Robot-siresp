from selenium import webdriver  
from selenium.webdriver.common.by import By 
from webdriver_manager.chrome import ChromeDriverManager  
from selenium.webdriver.chrome.service import Service  
import time  
from selenium.webdriver.support.ui import Select
import pandas as pd
from selenium.webdriver.common.by import By
from customtkinter import *
from tkinter import messagebox
import customtkinter as ctk
from CTkMessagebox import CTkMessagebox
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
import numpy as nan
from datetime import datetime
from customtkinter import *
import customtkinter as ctk
import json
from datetime import datetime


def salvar_historico(relatorio):
    try:
        # Carrega o histórico existente (se houver)
        with open('historico_agendamentos.json', 'r') as f:
            historico = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        historico = []
    
    # Adiciona os novos agendamentos ao histórico
    historico.extend(relatorio)
    
    # Salva o histórico atualizado
    with open('historico_agendamentos.json', 'w') as f:
        json.dump(historico, f, indent=4, ensure_ascii=False)


def carregar_historico():
    try:
        with open('historico_agendamentos.json', 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        return []
carregar_historico

relatorio = [] 
def agendado_sucesso():
    global agendamento_ok
    agendamento_ok = {
        'codigo' : indice['Código'],
        'profissional' : indice['Profissional'],
        'especialidade' : indice['Especialidade'], 
        'observação' : indice["Observação"],
        'status' : 'Agendado',
        'data' : datetime.now().strftime('%Y-%m-%d %H:%M:%S'),  # Adiciona a data do agendamento
    }
    relatorio.append(agendamento_ok)
    salvar_historico([agendamento_ok])  # Salva o agendamento no histórico
agendado_sucesso

def nao_agendado():
    global nao_agendamento
    nao_agendamento = {
        'Código' : indice['Código'],
        'profissional' : indice['Profissional'],
        'especialidade' : indice['Especialidade'], 
        'observação' : indice["Observação"],
        'status' : 'Não agendado',
    }
    relatorio.append(nao_agendamento)
    salvar_historico([agendamento_ok])
nao_agendado    

def gerar_relatorio_completo():
    historico = carregar_historico()
    if historico:
        df = pd.DataFrame(historico)
        nome_arquivo = f"Relatorio_Completo_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df.to_excel(nome_arquivo, index=False)
        CTkMessagebox(title="Sucesso!", message="Relatório completo gerado com sucesso!", icon="check", option_1="OK")
    else:
        CTkMessagebox(title="Aviso", message="Nenhum histórico encontrado!", icon="warning", option_1="OK")
gerar_relatorio_completo

def gerar_relatorio():
    if relatorio:
        df = pd.DataFrame(relatorio)
        relatorios_excel = f"Relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        df.to_excel(relatorios_excel, index=False)
        CTkMessagebox( title="Sucesso!", message="Relatorio Gerado com Sucesso!", icon="check", option_1="OK")
gerar_relatorio

def importar_planilha():
    global dicionario_pacientes
    caminho_pacientes = filedialog.askopenfilename(title="Selecione um arquivo Excel",filetypes=[("Excel", "*.xlsx")])  
    if caminho_pacientes:
        try:
            ctk.set_appearance_mode("blue")
            caminho_pacientes_leitura = pd.read_excel(caminho_pacientes)
            CTkMessagebox( title="Sucesso!", message="Operação concluída com sucesso!", icon="check", option_1="OK")
            caminho_pacientes_formatacao = caminho_pacientes_leitura[['Código', 'Especialidade', 'Profissional', 'Observação', 'Tipo Consulta']]
            dicionario_pacientes = caminho_pacientes_formatacao.to_dict('records')
            
        except Exception as e:
            messagebox.showerror("Erro", f" Não foi possível ler o arquivo:\n{e}")
importar_planilha

def iniciar_siresp():
    global drive_to_wait
    drive_to_wait = 10
    global navegador
    global iframe_site
    global iframe_principal
    global iframe_agenda_ce
    global iframe_frm
    servico = Service(ChromeDriverManager().install()) # Instala automaticamente o ChromeDriver
    navegador = webdriver.Chrome(service=servico)
    navegador.get("https://www.siresp.saude.sp.gov.br")
    botao_ambulatorio = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//*[@id="btn-4"]')))
    botao_ambulatorio.click()
    #navegador.find_element('xpath', '//*[@id="btn-4"]').click()
    if navegador:
        CTkMessagebox( title="Sucesso!", message="Navegador aberto com sucesso!", icon="check", option_1="OK")
iniciar_siresp

def final_planilha():
            verificacao = indice["codigo"]
            if pd.isna(verificacao):
                print('Não contém mais pacientes para ser agendados')
                navegador.close()
                CTkMessagebox(title="sucesso!",  message='Pacientes agendados com sucesso,\nAdicione uma nova planilha',icon="check", option_1="OK")
            return True   
final_planilha
        
def saindo_da_impress():
    try:
            alerta = WebDriverWait(navegador, drive_to_wait).until(ec.alert_is_present())
            #alerta = navegador.switch_to.alert
            alerta.dismiss()  # Clica em "Cancelar"
            print("Alerta cancelado")
            return True
    except:
        return False
saindo_da_impress

def Iniciar_robot(): 
    historico_existente = carregar_historico()
    print(f"Total de agendamentos no histórico: {len(historico_existente)}")
    global indice
    def retornando():
        navegador.switch_to.parent_frame()
        time.sleep(1)
        navegador.switch_to.parent_frame()
        caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO_CE"]')))
        dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO_CE"]')))
        navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
        select = Select(dropdown)
        tipo_consulta_desejada = 'Consulta'
        print('Tipo de consulta:', tipo_consulta_desejada)
        select_texto = select.select_by_visible_text(tipo_consulta_desejada)
        clickando = WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "FLT_F_COD_PACIENTE")]'))).click()
        limpando = WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "FLT_F_COD_PACIENTE")]'))).clear()
        #WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "Buscar")]'))).click()
    retornando
    
    def aba_agendamento():
        iframe_site = WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'site')))
        lista_menu = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.CLASS_NAME, 'sf-with-ul')))
        #lista_menu = navegador.find_elements('class name', 'sf-with-ul')
        for botao in lista_menu:
            if 'Agendamento' in botao.text:
                print("Agendamento Localizado")
                botao.click()
                # Localizando submenu de Agenda
                lista_agendamento = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//a[contains(@href, "age_fila.php?P=age_fila")]')))
                #lista_agendamento = navegador.find_elements('xpath', '//a[contains(@href, "age_marcar.php?P=age_marcar")]')
                print("Agenda Localizada")
                for botao_agenda in lista_agendamento:
                    print(botao_agenda.text)
                    if 'Cadastro Demanda por Recurso' in botao_agenda.text:
                        print("Elemento Localizado")
                        botao_agenda.click()
                    else:
                        print("Elemento não encontrado")
    aba_agendamento   
    
    aba_agendamento()
    
    def aba_listar():
        time.sleep(1)
        iframe_principal = WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'principal')))
        WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//td[contains(@id, "c_listar")]'))).click()
    aba_listar()

    for indice in dicionario_pacientes:
        
        def selecionar_especialidade():
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[contains(@name, "ID_ESP_EXA")]'))).click()
            caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@name, "ID_ESP_EXA")]')))
            dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@name, "ID_ESP_EXA")]')))
            navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            select = Select(dropdown)
            especialidade_desejada = indice['Especialidade']
            print(especialidade_desejada)
            select_texto = select.select_by_visible_text(especialidade_desejada)
            return
        selecionar_especialidade()
        time.sleep(1)
        def tipo_consulta():
            time.sleep(1)
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="TIPO"]'))).click()
            caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO"]')))
            dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO"]')))
            navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            select = Select(dropdown)
            tipo_consulta_desejada = indice['Tipo Consulta']
            print('Tipo de consulta:', tipo_consulta_desejada)
            select_texto = select.select_by_visible_text(tipo_consulta_desejada)
        tipo_consulta()
        
        def colocar_cod():
            time.sleep(1)
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "FLT_F_COD_PACIENTE")]'))).send_keys(indice['Código'])
        colocar_cod()

        def buscar():
            time.sleep(1)
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "Buscar")]'))).click()
        buscar()
        time.sleep(0.5)
        def selecionar_profissional():
            time.sleep(1)
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[contains(@name, "ID_PROFISSIONAL")]'))).click()
            caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@name, "ID_PROFISSIONAL")]')))
            dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@name, "ID_PROFISSIONAL")]')))
            navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            select = Select(dropdown)
            profissional_desejado = indice['Profissional']
            print(profissional_desejado)
            select_texto = select.select_by_visible_text(profissional_desejado)
            return
        selecionar_profissional()
        time.sleep(0.5)
        buscar()
        time.sleep(0.5)
        def selecionar_paciente():
           try:
                time.sleep(1)
                print("Formulário localizado")
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@type="radio" and @name="V_ID_FILA_SELECIONA"]'))).click()
                return True
           except:
               time.sleep(1)
               return False
        if not selecionar_paciente():
            retornando()
            CTkMessagebox(title="Erro!", message="PACIENTE JA POSSUI AGENDAMENTO NESTA ESPECIALIDADE!", icon="cancel", option_1="OK")
            agendado_sucesso()
            continue
    
        def sel_diaAg():
            try:
                global data_formatada
                global data_formatada_consulta
                iframe_agenda_ce = WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'agenda_ce')))
                time.sleep(2)
                print("Selecionando dia")
                div_mes0 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes0")]//td')))
                print("DIV 0 ENCONTRADA")
                div_mes1 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes1")]//td')))
                div_mes2 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes2")]//td')))
                div_mes3 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes3")]//td')))
                div_mes4 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes4")]//td')))
                try:
                    try:
                        for percorre in div_mes0:
                            mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes0')))
                            #mes = navegador.find_element(By.ID, 'mes0')
                            print(f"Percorrendo agenda de {mes.text}")
                            verificando = percorre.get_attribute("class")
                            data = percorre.get_attribute('id')
                            try:
                                data_formatada = (f"{data}").upper().split("|")[2]
                                print(data_formatada_consulta)
                                data_formatada_consulta = datetime.strptime(data_formatada,  "%Y-%m-%d").date()
                                print(data_formatada_consulta)
                            except: 
                                ...
                            print(verificando)
                            # Verifica se o texto contém a data desejada
                            condicao = "dia dados p" == verificando
                            condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                            if condicao or condicao2:
                                print(data)
                                print("Dia encontrado")
                                # Clica no dia desejado
                                percorre.click()
                                return True
                            else :
                                print("procurando próximo dia disponivel")
                        
                        WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes1'))).click()
                        #navegador.find_element(By.ID, 'mes1').click()   
                    except:
                        print('passando para o proximo paciente') 
                        return False
                    
                    # Percorrendo os dias do mês 1
                    try:
                        for percorre in div_mes1:
                            #time.sleep(0.5)
                            mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes1')))
                            #mes = navegador.find_element(By.ID, 'mes1')
                            print(f"Percorrendo agenda de {mes.text}")
                            verificando = percorre.get_attribute("class")
                            data = percorre.get_attribute('id')
                            print(data)
                            try:
                                data_formatada = (f"{data}").upper().split("|")[2]
                                print(data_formatada)
                                data_formatada_consulta = datetime.strptime(data_formatada,  "%Y-%m-%d").date()
                                print(data_formatada_consulta)
                            except: 
                                ...
                            print(verificando)
                            condicao = "dia dados p" == verificando
                            condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                            # Verifica se o texto contém a data desejada
                            if condicao or condicao2:
                                print(data_formatada_consulta)
                                print("Dia encontrado")
                                # Clica no dia desejado
                                percorre.click()
                                return True
                            else :
                                print("Dia não encontrado, procurando próximo dia")
                        
                        WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes2'))).click()       
                        #navegador.find_element(By.ID, 'mes2').click()
                        
                    except:
                        print('passando para o proximo paciente')
                        return False 
                    try:
                        for percorre in div_mes2:
                            mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes2')))
                            #mes = navegador.find_element(By.ID, 'mes2')
                            print(f"Percorrendo agenda de {mes.text}")
                            verificando = percorre.get_attribute("class")
                            data = percorre.get_attribute('id')
                            print(data)
                            try:
                                    data_formatada = (f"{data}").upper().split("|")[2]
                                    print(data_formatada)
                                    data_formatada_consulta = datetime.strptime(data_formatada,  "%Y-%m-%d").date()
                                    print(data_formatada_consulta)
                            except: 
                                 ...
                            print(verificando)
                            condicao = "dia dados p" == verificando
                            condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                            # Verifica se o texto contém a data desejada
                            if condicao or condicao2:
                                print(data)
                                print("Dia encontrado")
                                # Clica no dia desejado
                                percorre.click()
                                return True
                            else :
                                print("Dia não encontrado, procurando próximo")
                        WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes3'))).click()       
                    except:
                        print('passando para o proximo paciente')
                        return False   
                    
                    try:            
                        for percorre in div_mes3:
                            mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes3')))
                            #mes = navegador.find_element(By.ID, 'mes3')
                            print(f"Percorrendo agenda de {mes.text}")
                            verificando = percorre.get_attribute("class")
                            data = percorre.get_attribute('id')
                            print(data)
                            try:
                                data_formatada = (f"{data}").upper().split("|")[2]
                                print(data_formatada)
                                data_formatada_consulta = datetime.strptime(data_formatada,  "%Y-%m-%d").date()
                                print(data_formatada_consulta)
                            except: 
                                ...
                            print(verificando)
                            condicao = "dia dados p" == verificando
                            condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                            # Verifica se o texto contém a data desejada
                            if condicao or condicao2:
                                print(data)
                                print("Dia encontrado")
                                # Clica no dia desejado
                                percorre.click()
                                return True
                                break
                            else :
                                print("Dia não encontrado, procurando próximo")
                        WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes4'))).click()    
                        #navegador.find_element(By.ID, 'mes4').click()
                    except:
                        print('passando para o proximo paciente') 
                        return False
                    try:       
                        for percorre in div_mes4:
                            mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes4')))
                            #mes = navegador.find_element(By.ID, 'mes4')
                            print(f"Percorrendo agenda de {mes.text}")
                            verificando = percorre.get_attribute("class")
                            data = percorre.get_attribute('id')
                            try:
                                data_formatada = (f"{data}").upper().split("|")[2]
                                print(data_formatada)
                                data_formatada_consulta = datetime.strptime(data_formatada,  "%Y-%m-%d").date()
                                print(data_formatada_consulta)
                            except: 
                                ...
                            print(verificando)
                            condicao = "dia dados p" == verificando
                            condicao2 = "dia dados p azul_escuro p tp_interno" == verificando
                            # Verifica se o texto contém a data desejada
                            if condicao or condicao2:
                                print(data)
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
                return False
            except: 
                return False
        if not sel_diaAg():
            CTkMessagebox( title="Vagas", message="NÃO HÁ MAIS COTAS PARA ESTÁ ESPECIALIDADE!\nINSIRA UMA NOVA LISTA DE PACIENTES, DE UMA ESPECIALIDADE DIFERENTE", icon="check", option_1="OK")
            navegador.close()
            
        def cid():
            try:
                iframe_frm = WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'frm')))
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[contains(@class, "text")]'))).click()
                caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@class, "text")]')))
                dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[contains(@class, "text")]')))
                navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
                select = Select(dropdown)
                listaopc = select.options[0:]
                primopc = listaopc[1]
                if primopc:
                    primopc.click()
                    print(primopc.text)
                    return True
                else:
                    return
            except:
                print("Não precisa de CID!")
        cid()
        def selecionar_horario():
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="CB_CONSULTA"]'))).click()
            caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CB_CONSULTA"]')))
            dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CB_CONSULTA"]')))
            navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
            select = Select(dropdown)
            listaopc = select.options[0:]
            primopc = listaopc[1]
            if primopc:
                primopc.click()
                print(primopc.text)
                return True
            else:
                return
        selecionar_horario()
        time.sleep(0.5)
        def observacao():
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="AGE_OBSERVACAO"]'))).click()
            comentario = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//input[@name="AGE_OBSERVACAO"]')))
            comentario.send_keys(f"{indice['Observação']}------ROBOT")
        observacao()
        time.sleep(0.5)
        def botao_marcar():
            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="marcar_c"]'))).click()
        botao_marcar()
        
        def saindo_da_impress():
            alert = WebDriverWait(navegador, drive_to_wait).until(ec.alert_is_present())
            print(alert.text)
            alert.dismiss()
        saindo_da_impress()
        
        def alerta_fila():
            try:
                time.sleep(1)
                timeout = 3
                WebDriverWait(navegador, timeout).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="btn_fechar"]'))).click()
                return True
            except:
                print("Alerta não encontrado, ou conexão lenta")
                return False
        alerta_fila()
        
        # Volta para o contexto principal do navegador
        agendado_sucesso()
        time.sleep(1)
        retornando()
        time.sleep(1)

        # Passo 6: Verificando se o paciente foi agendado com sucesso
        COLETAR = 'COLETAR'
        COLETA = 'COLETA'
        observacao_FORMATADO = (f"{indice['Observação']}").upper()
        print(observacao_FORMATADO)
        if 'COLETA' in observacao_FORMATADO or'COLETAR' in observacao_FORMATADO:
            def sel_exames():
                time.sleep(1)
                print("SELECIONANDO EXAMES")
                caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO_CE"]')))
                dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="TIPO_CE"]')))
                navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
                select = Select(dropdown)
                tipo_consulta_desejada = 'Exame'
                print('Tipo de consulta:', tipo_consulta_desejada)
                select_texto = select.select_by_visible_text(tipo_consulta_desejada)
            sel_exames()
            
            time.sleep(1)
            
            def colocar_cod_exame():
                time.sleep(1)
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "FLT_F_COD_PACIENTE")]'))).send_keys(indice['Código'])
            colocar_cod_exame()
            
            time.sleep(1)
            
            def buscar_paciente_exames():
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "Buscar")]'))).click()
            buscar_paciente_exames()
            
            def selecionar_coleta():
                caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CBO_ID_GRUPO_COTA"]')))
                dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CBO_ID_GRUPO_COTA"]')))
                navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
                select = Select(dropdown)
                cota_desejada = 'LABORATÓRIO - INTERNO'
                print('Tipo de consulta:', cota_desejada)
                select_texto = select.select_by_visible_text(cota_desejada)
                time.sleep(1)
                #--------------------------------------------------------------
                caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="ID_ESP_EXA"]')))
                dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="ID_ESP_EXA"]')))
                navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
                select = Select(dropdown)
                exame_desejado = 'COLETA DE MATERIAL P/ EXAME LABORATORIAL'
                print('Tipo de consulta:', exame_desejado)
                select_texto = select.select_by_visible_text(exame_desejado)
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[contains(@name, "Buscar")]'))).click()
            selecionar_coleta()
            
            time.sleep(1.5)
            observacao_medica = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//td[@colspan ="20"]')))
            print(observacao_medica.text)
            obs_upper = observacao_medica.text.upper()
            obs_final = (f"{obs_upper}")
            profissional=(f"{indice["Profissional"]}").upper().split()
            print(obs_upper)
            time.sleep(2)
            for parte in profissional:
                if parte in obs_upper:
                    WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@type="radio" and @name="V_ID_FILA_SELECIONA"]'))).click()
                    
            def sel_dia_exa():
                try:
                    iframe_agenda_ce = WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'agenda_ce')))
                    time.sleep(2)
                    print("Selecionando dia")
                    div_mes0 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes0")]//td')))
                    print("DIV 0 ENCONTRADA")
                    div_mes1 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes1")]//td')))
                    div_mes2 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes2")]//td')))
                    div_mes3 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes3")]//td')))
                    div_mes4 = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_all_elements_located((By.XPATH, '//div[contains(@id, "div_mes4")]//td')))
                    try:
                        try:
                            for percorre in div_mes0:
                                mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes0')))
                                #mes = navegador.find_element(By.ID, 'mes0')
                                print(f"Percorrendo agenda de {mes.text}")
                                verificando = percorre.get_attribute("class")
                                data_exame = percorre.get_attribute('id')
                                try:
                                    data_formatada_exame = (f"{data_exame}").upper().split("|")[2]
                                    data_comparativa_exame = datetime.strptime(data_formatada_exame,  "%Y-%m-%d").date()
                                    print (data_comparativa_exame)
                                except: 
                                    ...
                                print(verificando)
                                # Verifica se o texto contém a data desejada
                                condicao = "dia dados p" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                condicao2 = "dia dados p azul_escuro p tp_interno" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                if condicao or condicao2:
                                    print(data_exame)
                                    print("Dia encontrado")
                                    # Clica no dia desejado
                                    percorre.click()
                                    return True
                                else :
                                    print("procurando próximo dia disponivel")
                            
                            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes1'))).click()
                            #navegador.find_element(By.ID, 'mes1').click()   
                        except:
                            print('passando para o proximo paciente') 
                            return False
                        
                        # Percorrendo os dias do mês 1
                        try:
                            for percorre in div_mes1:
                                #time.sleep(0.5)
                                mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes1')))
                                #mes = navegador.find_element(By.ID, 'mes1')
                                print(f"Percorrendo agenda de {mes.text}")
                                verificando = percorre.get_attribute("class")
                                data_exame = percorre.get_attribute('id')
                                try:
                                    data_formatada_exame = (f"{data_exame}").upper().split("|")[2]
                                    data_comparativa_exame = datetime.strptime(data_formatada_exame,  "%Y-%m-%d").date()
                                    print (data_comparativa_exame)
                                except: 
                                    ...
                                print(verificando)
                                condicao = "dia dados p" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                condicao2 = "dia dados p azul_escuro p tp_interno" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                # Verifica se o texto contém a data desejada
                                if condicao or condicao2:
                                    print(data_exame)
                                    print("Dia encontrado")
                                    # Clica no dia desejado
                                    percorre.click()
                                    return True
                                else :
                                    print("Dia não encontrado, procurando próximo dia")
                            
                            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes2'))).click()       
                            #navegador.find_element(By.ID, 'mes2').click()
                            
                        except:
                            print('passando para o proximo paciente')
                            return False 
                        try:
                            for percorre in div_mes2:
                                mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes2')))
                                #mes = navegador.find_element(By.ID, 'mes2')
                                print(f"Percorrendo agenda de {mes.text}")
                                verificando = percorre.get_attribute("class")
                                data_exame= percorre.get_attribute('id')
                                try:
                                    data_formatada_exame = (f"{data_exame}").upper().split("|")[2]
                                    data_comparativa_exame = datetime.strptime(data_formatada_exame,  "%Y-%m-%d").date()
                                    print (data_comparativa_exame)
                                except: 
                                    ...
                                print(verificando)
                                condicao = "dia dados p" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                condicao2 = "dia dados p azul_escuro p tp_interno" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                # Verifica se o texto contém a data desejada
                                if condicao or condicao2:
                                    print(data_exame)
                                    print("Dia encontrado")
                                    # Clica no dia desejado
                                    percorre.click()
                                    return True
                                else :
                                    print("Dia não encontrado, procurando próximo")
                            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes3'))).click()       
                        except:
                            print('passando para o proximo paciente')
                            return False   
                        
                        try:            
                            for percorre in div_mes3:
                                mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes3')))
                                #mes = navegador.find_element(By.ID, 'mes3')
                                print(f"Percorrendo agenda de {mes.text}")
                                verificando = percorre.get_attribute("class")
                                data_exame = percorre.get_attribute('id')
                                try:
                                    data_formatada_exame = (f"{data_exame}").upper().split("|")[2]
                                    data_comparativa_exame = datetime.strptime(data_formatada_exame,  "%Y-%m-%d").date()
                                    print (data_comparativa_exame)
                                except: 
                                    ...
                                print(verificando)
                                condicao = "dia dados p" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                condicao2 = "dia dados p azul_escuro p tp_interno" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                # Verifica se o texto contém a data desejada
                                if condicao or condicao2:
                                    print(data_exame)
                                    print("Dia encontrado")
                                    # Clica no dia desejado
                                    percorre.click()
                                    return True
                                    break
                                else :
                                    print("Dia não encontrado, procurando próximo")
                            WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.ID, 'mes4'))).click()    
                            #navegador.find_element(By.ID, 'mes4').click()
                        except:
                            print('passando para o proximo paciente') 
                            return False
                        try:       
                            for percorre in div_mes4:
                                mes = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.ID, 'mes4')))
                                #mes = navegador.find_element(By.ID, 'mes4')
                                print(f"Percorrendo agenda de {mes.text}")
                                verificando = percorre.get_attribute("class")
                                data_exame = percorre.get_attribute('id')
                                try:
                                    data_formatada_exame = (f"{data_exame}").upper().split("|")[2]
                                    data_comparativa_exame = datetime.strptime(data_formatada_exame,  "%Y-%m-%d").date()
                                    diferenca = abs((data_formatada_consulta - data_comparativa_exame).days)
                                    print (data_comparativa_exame)
                                except: 
                                    ...
                                    
                                print(verificando)
                                condicao = "dia dados p" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                condicao2 = "dia dados p azul_escuro p tp_interno" == verificando and 15 <= abs((data_formatada_consulta - data_comparativa_exame).days) <= 20
                                # Verifica se o texto contém a data desejada
                                if condicao or condicao2:
                                    print(data_exame)
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
                    return False
                except: 
                    return False
            if not sel_dia_exa():
                CTkMessagebox( title="Vagas", message="NÃO HÁ MAIS COTAS PARA ESTÁ ESPECIALIDADE!\nINSIRA UMA NOVA LISTA DE PACIENTES, DE UMA ESPECIALIDADE DIFERENTE", icon="check", option_1="OK")
                navegador.close()
                
            def selecionar_horario_exame():
                WebDriverWait(navegador, drive_to_wait).until(ec.frame_to_be_available_and_switch_to_it((By.ID, 'frm')))
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//select[@name="CB_EXAME"]'))).click()
                caminho = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CB_EXAME"]')))
                dropdown = WebDriverWait(caminho, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//select[@name="CB_EXAME"]')))
                navegador.execute_script("arguments[0].scrollIntoView(true);", dropdown)
                select = Select(dropdown)
                listaopc = select.options[0:]
                primopc = listaopc[1]
                if primopc:
                    primopc.click()
                    print(primopc.text)
                    return True
                else:
                    return
            selecionar_horario_exame()
        
            def observacao_exame():
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="AGE_OBSERVACAO"]'))).click()
                comentario = WebDriverWait(navegador, drive_to_wait).until(ec.presence_of_element_located((By.XPATH, '//input[@name="AGE_OBSERVACAO"]')))
                comentario.send_keys(f"A PDD DO DR/DRA {indice['Profissional']}------ROBOT")
            observacao_exame()
            
            time.sleep(0.5)
            
            def marcar_exame():
                WebDriverWait(navegador, drive_to_wait).until(ec.element_to_be_clickable((By.XPATH, '//input[@name="marcar_e"]'))).click()
            marcar_exame()
            
            saindo_da_impress()
            alerta_fila()
            retornando()
    navegador.close()
    CTkMessagebox( title="Agendado!", message="AGENDADO COM SUCESSO!", icon="check", option_1="OK")

Iniciar_robot
