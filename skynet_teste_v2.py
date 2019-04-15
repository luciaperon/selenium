import time
import pandas as pd
import tkinter
import re
from datetime import datetime
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.common.keys import Keys

############################# LER INPUTS
# cursos elegiveis
df = pd.DataFrame()

for i in ['Olimpo', 'Siae', 'Colaborar']:
    df_aux = pd.read_excel('cursos_suporte_DNM_19.xlsx', sheet_name=i)
    df = df.append(df_aux)

df['chave_elegivel'] = df['curso'] + df['semestre'].astype(str)
lista_cursos_elegiveis = list(df['chave_elegivel'])

# ler status aluno
df_status_aluno = pd.read_csv('//10.100.63.238/Departamentos/GAA-ENADE/Desafio Nota Máxima/Status Aluno/status_aluno.csv', engine='python')
lista_cpfs_dnm = list(df_status_aluno['user_cpf'])


######################################### MENSAGENS
msg1 = """E aí {},
Consultamos no sistema e vimos que você já possui acesso.
Acesse o DNM com o seu cpf nos campos login e senha ou clique em esqueci a minha senha, que mandamos uma senha nova para o seu e-mail!
Durante o login, ele atualiza a página e mostra algumas opções de semestres para logar, sempre de uma olhada se está em 2019.1.
Bons estudos!"""#### OK ####

msg2 = """E aí {},
Consultamos no sistema e não encontramos o seu cpf na base.
Verifique se você digitou corretamente o número do cpf, caso esteja, procure seu coordenador pois não te encontramos vinculado no sistema.
Bons estudos!"""#### ERRO ####

msg3 = """E aí {},
Checamos aqui e você realmente é um aluno DNM, estamos te devendo.
Nós atualizamos diariamente os usuários, verifique dentro de 24h se você conseguirá acessar!
Lembre-se: Para acessar o dnm use o seu cpf nos campos login e senha, sempre de uma olhada se está em 2019.1.
Bons estudos!"""

msg4 = """E aí {},
Infelizmente seu curso e semestre não participam do DNM 2019.1.
Se ficar com alguma dúvida do porque, procure alguém na sua unidade para lhe explicar como funciona a regra.
Estamos ansiosos para te receber nos próximos!
Bons estudos!"""#### ERRO ####

msg5 = """E aí {},
Verificamos que você é um aluno elegível, mas não está matriculado na turma da disciplina participante do DNM.
Confirme com seu coordenador se está tudo certo com suas disciplinas e turma, que ele poderá lhe ajudar em algum problema mais complexo.
Bons estudos!"""

msg6 = """E aí {},
Verificamos que você não enviou o CPF, favor abrir um novo chamado, informando seu CPF.
Bons estudos!"""

# elementos para se movimentar nas paginas
elem_responder = '//*[@id="mail_message_form"]/div/div/div/button'
elem_ok = '/html/body/div[5]/div/div/div[2]/button'
elem_ok2 = '/html/body/div[6]/div/div/div[2]/button'
elem_analise = '//*[@id="mail_message_form"]/div/div/div/ul/li[1]/a'
pagina_suporte = "http://127.0.0.1:8000/signup/"

























def abrir_dnm():
    driver = webdriver.Chrome(executable_path=r"chromedriver.exe")
    #driver = webdriver.Firefox()
    driver.maximize_window()
    driver.get("https://www.desafionotamaxima.com.br/login?locale=pt-BR")

    time.sleep(5)

    ######################## LOGAR NO DNM E IR PRA PAGINA DE CHAMADO
    # usuario do Thalles
    username = driver.find_element_by_id("nome").send_keys("123456789000")
    password = driver.find_element_by_id("senha").send_keys("123")

    element = driver.find_element_by_xpath('//*[@id="new_user_session"]/fieldset[4]/input').click()

    """
    ###################################### abrir pagina de chamado
    element = driver.find_element_by_xpath('//*[@id="js-dtour-step-two"]')
    element.click()

    element = driver.find_element_by_xpath('//*[@id="js-new-design-tour"]/div[2]/div/ul/li[2]/ul/li[1]/a')
    element.click()
    """

    # enquanto o link ta quebrado
    driver.get("https://www.desafionotamaxima.com.br/mail_messages?locale=pt-BR")

    return driver


def abrir_suporte():
    ############################ ABRIR PAGINAS
    driver2 = webdriver.Chrome(executable_path=r"chromedriver.exe")
    #driver = webdriver.Firefox()
    driver2.maximize_window()
    driver2.get(pagina_suporte)

    return driver2


def atualizar_chamados():

    # filtrar por status
    elem_filtro1 = '//*[@id="accordion1"]/div/div[1]/h4'
    elem_aberto = '//*[@id="collapseStatus"]/div/div/div[1]/div[2]/*[@id="array_status_"]'
    print('a1')
    result = None
    while result is None:
        try:
            print('a2')
            driver.get("https://www.desafionotamaxima.com.br/mail_messages?locale=pt-BR")
            time.sleep(10)
            element = driver.find_element_by_xpath(elem_filtro1)
            result = 1
        except:
            print('a3')
            pass



    ###################################### filtrar chamados
    time.sleep(0.5)
    element = driver.find_element_by_xpath(elem_filtro1).click()
    print('a4')
    ###################################### esperar abrir o filtro
    result = None
    while result is None:
        try:
            print('a5')
            # connect
            element = driver.find_element_by_xpath(elem_aberto).click()
            result = 1
        except:
            print('a6')
            pass

    time.sleep(0.5)
    element = driver.find_element_by_xpath(elem_filtro1).click()

    ###################################### filtrar por motivo detalhado
    elem_filtro2 = '//*[@id="accordion2"]/div/div[1]/h4/a'
    elem_dificuldade_acesso = '//*[@id="collapseSMotivo"]/div/div/div[3]/div[5]//*[@id="array_subtype_"]'

    time.sleep(0.5)
    element = driver.find_element_by_xpath(elem_filtro2).click()
    print('a7')
    # esperar abrir o filtro
    result = None
    while result is None:
        try:
            print('a8')
            # connect
            element = driver.find_element_by_xpath(elem_dificuldade_acesso).click()
            result = 1
        except:
            print('a9')
            pass

    time.sleep(3)
    element = driver.find_element_by_xpath(elem_filtro2).click()

    ###################################### filtrar alunos
    elem_filtro_alunos = '/html/body/div[3]/div[2]/div/form/div[4]/div[2]/div/div[2]/a[1]'
    time.sleep(1)
    print('a10')
    result = None
    while result is None:
        try:
            print('a11')
            # connect
            element = driver.find_element_by_xpath(elem_filtro_alunos).click()
            result = 1
        except:
            print('a12')
            pass

    time.sleep(3)

def verificar_chamado_abertos():
    time.sleep(5)
    elem_chamado = '//*[@id="call_btn_0"]'
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_chamado)
            result = 1
            status_data = 'com_chamados'
        except:
            result = 1
            status_data = 'sem_chamados'

    return status_data



def ordenar_chamados_abrir_primeiro():
    ####################### ordenar por ordem de criação e esperar abrir o filtro
    elem_data_criacao = '/html/body/div[3]/div[2]/div/form/div[4]/div[3]/table/thead/tr/th[7]/a'
    print('a13')
    result = None
    while result is None:
        try:
            print('a14')
            # connect
            element = driver.find_element_by_xpath(elem_data_criacao).click()
            result = 1
        except:
            print('a15')
            pass

    ###################################### primeiro chamado para resposta
    elem_chamado = '//*[@id="call_btn_0"]'
    time.sleep(3)
    element = driver.find_element_by_xpath(elem_chamado).click()




def apertar_ok():
    ############################ apertar ok pra fechar chamado
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_ok).click()
            result = 1
        except:
            pass

        try:
            # connect
            element = driver.find_element_by_xpath(elem_ok2).click()
            result = 1
        except:
            pass

def mandar_pra_analise():
    element = driver.find_element_by_xpath(elem_responder).click()
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_analise).click()
            result = 1
        except:
            pass

    apertar_ok()

def verificar_load(elem_verificar):
    result = None
    while result is None:
        try:
            # connect
            element = driver.find_element_by_xpath(elem_verificar).click()
            result = 1
        except:
            pass

def esperar_pagina():
    elem_controller = '//*[@id="full-page-wrapper"]/div[2]/div/div[1]/div/div[1]/h3'
    elem_controller2 = '//*[@id="js-dtour-step-one"]/div/div[2]/ul/li[5]/a'
    result = None
    while result is None:
        try:
            try:
                element = driver.find_element_by_xpath(elem_controller).text
                result = 1
            except:
                element = driver.find_element_by_xpath(elem_controller2).text
                result = 1
        except:
            pass

def verificar_chamado():
    try:
        element = driver.find_element_by_xpath(elem_responder).text
        return True
    except:
        return False

def responder_chamado():
    ##################################### verificar data
    elem_data_criacao = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[4]'
    elem_data_modificacao = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[5]'

    # verificar se carregou
    verificar_load(elem_data_criacao)

    data_criacao = driver.find_element_by_xpath(elem_data_criacao).text
    data_criacao = data_criacao.replace('Data de criação\n\n', '')

    data_modificacao = driver.find_element_by_xpath(elem_data_modificacao).text
    data_modificacao = data_modificacao.replace('Data de última modificação\n\n', '')

    ################################## VERIFICAR SEM ACESSO
    elem_aluno = '//*[@id="sidebar_adaptativos"]/div/div/div/fieldset[3]'
    aluno = driver.find_element_by_xpath(elem_aluno).text
    aluno = aluno.replace('Aluno\n\n', '')

    print(data_criacao, aluno)

    if aluno == 'SEM ACESSO':
        try:
            ##################################### SALVAR NOME DO ALUNO
            elem_nome = '//*[@id="full-page-wrapper"]/div[{}]/div/div[3]/div[2]/div/p[1]'

            try:
                nome = driver.find_element_by_xpath(elem_nome.format(2)).text

            except:
                nome = driver.find_element_by_xpath(elem_nome.format(3)).text


            nome = nome.replace('Nome: ', '')
            nome = nome.split()[0].title()

            ###################################### PEGAR CPF
            elem_cpf = '//*[@id="full-page-wrapper"]/div[{}]/div/div[3]/div[2]/div/p[3]'
            time.sleep(2)
            try:
                cpf = driver.find_element_by_xpath(elem_cpf.format(2)).text

            except:
                cpf = driver.find_element_by_xpath(elem_cpf.format(3)).text

            cpf = re.sub('[^0-9]', '', cpf)

            try:
                cpf_check_dnm = int(cpf)

            except:
                pass

            if cpf == '':
                mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg6.format(nome))

            ################################## ver se o cpf ja esta no DNM
            elif cpf_check_dnm in lista_cpfs_dnm:
                mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg1.format(nome))

            else:
                ###################################### BUSCAR CPF no SUPORTE
                username = driver2.find_element_by_xpath('//*[@id="id_cpf"]').send_keys(cpf)
                element = driver2.find_element_by_xpath('/html/body/div/div/div/form/center/input').click()

                ####################### VERIFICAR CPF NO SERVIDOR DO SUPORTE
                elem_primeira_linha_suporte = '/html/body/table/tbody/tr/th'
                elem_head_cpf = '/html/body/table/thead/tr/th[2]'

                ######################### VERIFICAR SE JA CARREGOU
                try:
                    # esperar aparecer a base do suporte
                    result = None
                    while result is None:
                        try:
                            # connect
                            driver2.find_element_by_xpath(elem_head_cpf).text
                            result = 1
                        except:
                            driver2.refresh()

                    driver2.find_element_by_xpath(elem_primeira_linha_suporte).text

                    if 'DNM' in driver2.page_source:
                        ############## ALUNO TEM DISCIPLINA VINCULADA
                        mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg3.format(nome))
                    else:
                        ################### VERIFICAR SE É ELEGIVEL
                        elem_curso = '/html/body/table/tbody/tr[1]/td[9]'
                        elem_semestre = '/html/body/table/tbody/tr[1]/td[10]'
                        curso = driver2.find_element_by_xpath(elem_curso).text
                        semestre = driver2.find_element_by_xpath(elem_semestre).text

                        tag_elegivel = curso + semestre

                        if tag_elegivel in lista_cursos_elegiveis: # trazer curso e semestre e verificar se esta no excel
                            mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg5.format(nome))
                        else:
                            mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg4.format(nome))

                except:
                    ############################# NAO VEIO TABELA NO SUPORTE
                    mensagem = driver.find_element_by_id("mail_message_content").send_keys(msg2.format(nome))


            time.sleep(0.5)
            ######################################### clicar em respondido
            element = driver.find_element_by_xpath(elem_responder).click()
            elem_respondido = '//*[@id="mail_message_form"]/div/div/div/ul/li[3]/a'

            result = None
            while result is None:
                try:
                    # connect
                    element = driver.find_element_by_xpath(elem_respondido).click()
                    result = 1
                except:
                    pass

            apertar_ok()
            print('respondido')

        except:
            mandar_pra_analise()
            print('Análise pq deu erro')
            return 'g'


    else:
        mandar_pra_analise()
        print('Análise pq é aluno')
















while True:
    debug = 'adas'

    # pegar hora
    n = datetime.now()
    t = n.timetuple()
    y, m, d, h, min, sec, wd, yd, i = t

    # ler o status aluno as 10
    if h == 10 and min == 0 and sec == 0:
        try:
            # status aluno
            df_status_aluno = pd.read_csv('Z:/GAA-ENADE/Desafio Nota Máxima/Status Aluno/status_aluno.csv', engine='python')
            lista_cpfs_dnm = list(df_status_aluno['user_cpf'])
        except:
            pass

    # responder chamados a cada hora
    elif min == 10 and sec == 0:
        print('a')
        driver = abrir_dnm()
        time.sleep(6)
        driver2 = abrir_suporte()
        time.sleep(3)
        atualizar_chamados()

        status_data = verificar_chamado_abertos()

        if status_data == 'sem_chamados':
            pass
        else:
            ordenar_chamados_abrir_primeiro()

            result = None
            while result is None:
                print('b')
                esperar_pagina()
                time.sleep(2)
                print('c')
                if verificar_chamado() is True:
                    print('d')
                    debug = responder_chamado()
                    driver2.get(pagina_suporte)
                    print('e')
                else:
                    print('f')
                    result = 1

        driver.quit()
        driver2.quit()
