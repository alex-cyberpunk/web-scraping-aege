# -*- coding: utf-8 -*-
"""
Created on Tue May  3 11:25:05 2022

@author: alex.matias
"""
import threading
import time
import os
import pandas 
from selenium import webdriver
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select


#######################entrada#############################
path=r'web-scraping-aege\planilhas de entrada'# le os templates de entrada
nome_parque=''#para o salvamento de screenshots da saida
#########################################
def retorna_lista_entrada(path,aba,coluna):
    lista=[]
    excel_data_df = pandas.read_excel(path, sheet_name=aba,header=0)
    #print(excel_data_df)
    if coluna in excel_data_df.columns:
        lista=(excel_data_df[coluna].tolist())
    #print(lista)
    return lista

def le_arquivo_excel(path,nome_planilha,aba,coluna):
    lista_entrada=[]
    dirs = os.listdir(path)
    for file in dirs:
            if file.endswith(".xlsx"):
                #print(file)
                if file == nome_planilha+str('.xlsx'):
                    diretorioParques1 = os.path.join(path, file)
                    #print(diretorioParques1)
                    lista_entrada=retorna_lista_entrada(str(diretorioParques1),aba,coluna)
    return lista_entrada

def covert_string(lista):
    temp=[]
    for i in range(len(lista)):
        temp2=str(lista[i]).replace('.', ',')
        temp.append(temp2)
    return temp
def convert_string_sem_numeral(lista):
    temp=[]
    for i in range(len(lista)):
        temp2=str(lista[i])
        temp.append(temp2)
    return temp
def convert_string_sem_nan(lista):
    temp=[]
    for i in range(len(lista)):
        temp2=str(lista[i])
        if temp2!='nan':
            temp.append(temp2)
    return temp
def espera_alert(browser):
        WebDriverWait(browser, 100).until(EC.alertIsPresent(),'Timed out waiting for PA creation ' +'confirmation popup to appear.')
        alert = browser.switch_to.alert
        alert.accept()
        print("alert accepted")
def encontra_texto_no_site(num_rows,num_cols,browser,parque):
    before_XPath = "//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr["
    aftertd_XPath = "]/td["
    aftertr_XPath = "]"
    for i in range(len(parque)):
        for t_row in range(2, (num_rows + 1)):
                for t_column in range(1, (num_cols + 1)):
                    FinalXPath = before_XPath + str(t_row) + aftertd_XPath + str(t_column) + aftertr_XPath
                    cell_text =browser.find_element(By.XPATH, FinalXPath).text
                    #cell_text = browser.find_element_by_xpath(FinalXPath).text
                    if(cell_text==parque):return FinalXPath
def exclui_itens_lista_site(browser,colunas):
    editar='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Editar"]'
    excluir='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Excluir"]'
    salvar='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Salvar"]'
    #encontra_texto_adaptado
    num_rows = 2
    num_cols = colunas
    before_XPath = "//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr["
    aftertd_XPath = "]/td["
    aftertr_XPath = "]"
    time.sleep(5)
    if(num_rows!=0):
        for t_row in range(2, (num_rows + 1)):
            for t_column in range(1, (num_cols + 1)):
                FinalXPath = before_XPath + str(t_row) + aftertd_XPath + str(t_column) + aftertr_XPath
                cell_text =browser.find_element(By.XPATH, FinalXPath).text
                if(cell_text!=""):
                    wait.until(EC.element_to_be_clickable((By.XPATH, excluir))).click()
                    time.sleep(12)
                    #browser.switchTo().alert().accept()
        #wait.until(EC.element_to_be_clickable((By.XPATH, salvar))).click()
def login(login,senha):
    browser= webdriver.Chrome(ChromeDriverManager().install())
    browser.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
        "source":
            "const newProto = navigator.__proto__;"
            "delete newProto.webdriver;"
            "navigator.__proto__ = newProto;"
        })
    browser.get('https://aege-empreendedor.epe.gov.br')
    time.sleep(10)
    username = browser.find_element(By.NAME,'sph_username')
    #username = browser.find_element_by_name('sph_username')
    password = browser.find_element(By.NAME,'sph_password')
    #password = browser.find_element_by_name('sph_password')
    username.send_keys(login)
    password.send_keys(senha)
    browser.find_element(by=By.XPATH,value='/html/body/div/div/form/div[4]/input').click() 
    return browser
def retorna_aba(nome):
    if(nome=='Capacidade'): aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabLeilao"]'
    if(nome=='Caracteristicas tecnicas'):aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabCaracteristicasTecnicas"]'
    if(nome=='Equipamentos'):aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabEquipamentos"]'
    if(nome=='Outorgas'):aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabOutorgas"]'
    if(nome=='Leilao'): aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabLeilao"]'
    if(nome=='Caracteristicas tecnicas'):aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabCaracteristicasTecnicas"]'
    if(nome=='Conexao'):aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabConexao"]'
    return aba
    
def retorna_sub_aba(nome):
    if(nome=='Torres'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabTorresMedicao"]'
    if(nome=='Local'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabDadosLocal"]'
    if(nome=='Dados Anenométricos'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabDadosAnemometricosCertificados"]'
    if(nome=='Rosa'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabRosaDosVentos"]'
    if(nome=='Prod'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabProducaoEnergia"]'
    if(nome=='Prod-Aero'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabProducaoAerogerador"]'
    if(nome=='Coordenadas'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabCoordenadaAerogerador"]'
    if(nome=='Aeros'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabAerogerador"]'
    if(nome=='Cronograma'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabCronograma"]'
    if(nome=='Orcamento'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabOrcamento"]'
    if(nome=='Moto'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabMotorizacao"]'
    if(nome=='REIDI'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabREIDI"]'
    if(nome=='Informação'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabInfoEnergeticas"]'
    if(nome=='Subestação'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabSubestacao"]'
    if(nome=='Instalação'):sub_aba='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabLinhaTransmissao"]'
    return sub_aba
def retorna_elementos_da_sub_aba(nome):
    if(nome=='Torres'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalCoordenadaUTMNTorre"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalCoordenadaUTMETorre"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalFusoUTMTorre"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade01"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade02"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade03"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade04"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade05"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade06"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade07"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade08"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade09"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade10"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade11"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade12"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade13"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade14"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade15"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade16"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade17"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade18"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade19"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade20"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade21"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade22"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade23"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade24"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVelocidade25"]']
    if(nome=='Dados do Local'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalAltitudeMedia"]','//*[@id="ctl00_ContentPlaceHolder1_colvalTemperaturaMedia"]','//*[@id="ctl00_ContentPlaceHolder1_colprcUmidadeRelativaMediaAnual"]','//*[@id="ctl00_ContentPlaceHolder1_colvalPressaoAtmosferica"]','//*[@id="ctl00_ContentPlaceHolder1_colvalRugosidadeMediaTerreno"]']
    if(nome=='Dados Anenométricos'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalRugosidadeMediaTerreno"]','//*[@id="ctl00_ContentPlaceHolder1_coldatIniReferenciaDA"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimReferenciaDA"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_01"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_02"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_03"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_04"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_05"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_06"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_07"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_08"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_09"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_10"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_11"]','//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeVento_0_12"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_01"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_02"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_03"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_04"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_05"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_06"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_07"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_08"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_09"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_10"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_11"]','//*[@id="ctl00_ContentPlaceHolder1_colvalDensidadeAr_0_12"]']
    if(nome=='Rosa'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento1"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento2"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento3"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento4"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento5"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento6"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento7"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento8"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento9"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento10"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento11"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento12"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento13"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento14"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento15"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcFrequenciaVento16"]']
    if(nome=='Prod'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_01"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_02"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_03"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_04"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_05"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_06"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_07"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_08"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_09"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_10"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_11"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_12"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoCertificada_0_13"]']
    if(nome=='Prod-Aero'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalVelocidadeAnualVentoLivreTurbina"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalProducaoAnualEnergiaBruta"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcPerdaAerodinamica"]', '//*[@id="ctl00_ContentPlaceHolder1_colprcDegradacaoMediaPas"]', '//*[@id="ctl00_ContentPlaceHolder1_colcodTorreMedicaoReferencia"]']
    if(nome=='Outorgas'):lista=['//*[@id="ctl00_ContentPlaceHolder1_coltxtOrgaoEmissor"]', '//*[@id="ctl00_ContentPlaceHolder1_colnumLicenca"]', '//*[@id="ctl00_ContentPlaceHolder1_coldatEmissao"]', '//*[@id="ctl00_ContentPlaceHolder1_coldatValidade"]', '//*[@id="ctl00_ContentPlaceHolder1_colcodModalidade"]', '//*[@id="ctl00_ContentPlaceHolder1_colindContemplaConexao_1"]', '//*[@id="ctl00_ContentPlaceHolder1_colnomLogradouroCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_colnomBairroCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_colsigUFCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_colnumEnderecoCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_colnumCEPCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_coltxtComplementoCorrespondencia"]', '//*[@id="ctl00_ContentPlaceHolder1_colnomMunicipioCorrespondencia"]']
    if(nome=='Capacidade'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colindEmpreendimentoNovo"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalPotenciaInstalada"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalRepotenciacao"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalAmpliacao"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalPotenciaTotalInstalada"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalPotenciaInjetavelMax"]']
    if(nome=='Coordenadas'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colnomAerogerador"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalCoordenadaUTMEAero"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalCoordenadaUTMNAero"]', '//*[@id="ctl00_ContentPlaceHolder1_colindHemisferioFusoUTMAero_1"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalFusoUTMAero"]']
    if(nome=='Aeros'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colcodTipoTurbina"]', '//*[@id="ctl00_ContentPlaceHolder1_colcodTipoClasseTurbina"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalAlturaEixo"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalPotenciaMaximaGerada"]']
    if(nome=='Orcamento'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalAquisicaoTerreno"]','//*[@id="ctl00_ContentPlaceHolder1_colvalMeioAmbiente"]','//*[@id="ctl00_ContentPlaceHolder1_colvalObraCivil"]','//*[@id="ctl00_ContentPlaceHolder1_colvalEquipamento"]','//*[@id="ctl00_ContentPlaceHolder1_colvalSistemaComunicacao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalTransporteSeguro"]','//*[@id="ctl00_ContentPlaceHolder1_colvalMontagemTeste"]', '//*[@id="ctl00_ContentPlaceHolder1_colvalCustoIndireto"]','//*[@id="ctl00_ContentPlaceHolder1_colvalAquisicaoTerrenoConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalMeioAmbienteConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalObraCivilConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalLinhaTransmissaoConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalTransporteSeguroConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalMontagemTesteConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalCustoIndiretoConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colvalSubestacao"]']
    if(nome=='Cronograma'):lista=['//*[@id="ctl00_ContentPlaceHolder1_coldatInicioObra"]','//*[@id="ctl00_ContentPlaceHolder1_coldatComprovacaoAporte"]','//*[@id="ctl00_ContentPlaceHolder1_coldatObtencaoLI"]','//*[@id="ctl00_ContentPlaceHolder1_coldatObtencaoLO"]','//*[@id="ctl00_ContentPlaceHolder1_coldatInicioCanteiroObra"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimCanteiroObra"]','//*[@id="ctl00_ContentPlaceHolder1_colprzTotal"]','//*[@id="ctl00_ContentPlaceHolder1_coldatComprovacaoContratoFornecimentoEquipamento"]','//*[@id="ctl00_ContentPlaceHolder1_coldatInicioObraCivilEstrutura"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimObraCivilEstrutura"]','//*[@id="ctl00_ContentPlaceHolder1_coldatInicioConcretagemBase"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimConcretagemBase"]','//*[@id="ctl00_ContentPlaceHolder1_coldatInicioMontagemTorre"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimMontagemTorre"]','//*[@id="ctl00_ContentPlaceHolder1_coldatInicioObraSubestacao"]','//*[@id="ctl00_ContentPlaceHolder1_coldatFimObraSubestacao"]']
    if(nome=='REIDI'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalBensComImposto"]','//*[@id="ctl00_ContentPlaceHolder1_colvalServicosComImposto"]','//*[@id="ctl00_ContentPlaceHolder1_colvalOutrosInvestimentosComImposto"]','//*[@id="ctl00_ContentPlaceHolder1_colqtdEmpregoDiretoGerado"]','//*[@id="ctl00_ContentPlaceHolder1_colvalBensSemImposto"]','//*[@id="ctl00_ContentPlaceHolder1_colvalServicosSemImposto"]','//*[@id="ctl00_ContentPlaceHolder1_colvalOutrosInvestimentosSemImposto"]']
    if(nome=='Informação'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colvalConsumoProprio"]']
    if(nome=='Subestação'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colindTipoSubestacaoSESE"]','//*[@id="ctl00_ContentPlaceHolder1_colcodTipoTensaoSuperiorSESE"]','//*[@id="ctl00_ContentPlaceHolder1_colqtdTrafoSubestacaoSETE"]','//*[@id="ctl00_ContentPlaceHolder1_colcodTensaoEnrolamentoPrimarioSubestacaoSETE"]','//*[@id="ctl00_ContentPlaceHolder1_colvalPotenciaNominalTrafoSubestacaoSETE"]','//*[@id="ctl00_ContentPlaceHolder1_colcodTensaoEnrolamentoSecundarioSubestacaoSETE"]']
    if(nome=='Instalação'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colindTipoConexaoIC_0"]','//*[@id="ctl00_ContentPlaceHolder1_colcodPontoConexao"]','//*[@id="ctl00_ContentPlaceHolder1_colindIgualdadeNivelTensaoSE_0"]','//*[@id="ctl00_ContentPlaceHolder1_colqtdCircuitoLTE"]','//*[@id="ctl00_ContentPlaceHolder1_colindTipoCircuitoLTE_0"]','//*[@id="ctl00_ContentPlaceHolder1_colBitolasLTE"]','//*[@id="ctl00_ContentPlaceHolder1_colvalExtensaoLTE"]']
    if(nome=='Moto'):lista=['//*[@id="ctl00_ContentPlaceHolder1_colnumUnidadeGeradora"]','//*[@id="ctl00_ContentPlaceHolder1_colvalPotencia"]','//*[@id="ctl00_ContentPlaceHolder1_coldatOperacaoTeste"]','//*[@id="ctl00_ContentPlaceHolder1_coldatOperacao"]']
    return lista

def tipo_de_entrada(sub_aba_nome,posicao):
    tipo='input'
    if (sub_aba_nome=='Prod-Aero'):
        if(posicao==4):
            tipo='select'
    elif (sub_aba_nome=='Outorgas'):
        if(posicao==4):
            tipo='select'
    elif (sub_aba_nome=='Outorgas'):
        if(posicao==8):
            tipo='select'
    elif (sub_aba_nome=='Aeros'):
        if(posicao==0 or posicao==1):
            tipo='select'
    elif (sub_aba_nome=='Capacidade'):
        if(posicao==0):
            tipo='select'
    elif (sub_aba_nome=='Capacidade'):
        if(posicao==0):
            tipo='click'
    elif (sub_aba_nome=='Subestação'):
        if(posicao==0 or posicao==1 or posicao==3 or posicao==5):
            tipo='select'
    elif (sub_aba_nome=='Instalação'):
        if(posicao==1 or posicao==3 or posicao==5):
            tipo='select'
        if(posicao==0 or posicao==2 or posicao==4):
            tipo='click'
    elif (sub_aba_nome=='Coordenadas'):
        if(posicao==3):
            tipo='click'
    return tipo
def preenchimento(browser,wait,aba_nome,sub_aba_nome,itens_p):
    editar='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Editar"]'
    incluir='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Novo"]'
    aba=retorna_aba(aba_nome)
    wait.until(EC.element_to_be_clickable((By.XPATH, aba)))
    wait.until(EC.element_to_be_clickable((By.XPATH, aba))).click()
    print(aba_nome)
    time.sleep(15)
    
    if(sub_aba_nome=='Moto_s'):sub_aba=retorna_sub_aba('Moto')
    else:sub_aba=retorna_sub_aba(sub_aba_nome)
    
    
    wait.until(EC.element_to_be_clickable((By.XPATH, sub_aba)))
    wait.until(EC.element_to_be_clickable((By.XPATH, sub_aba))).click()
    print(sub_aba_nome)
    time.sleep(12) 
    #valor=input('tamanho onde parou : ')
    if(sub_aba_nome=='Coordenadas'):
        num_colunas=int(input('quantos aeros tem?:(numero inteiro)'))
        exclui_itens_lista_site(browser,num_colunas)
        tamanho=len(itens_p)
    elif(sub_aba_nome=='Prod-Aero'):
        tamanho=len(itens_p)
    elif(sub_aba_nome=='Moto'):tamanho=int(itens_p[0])
    else:tamanho=1
    for j in range(tamanho):
        if(sub_aba_nome=='Coordenadas'):
            itens=itens_p[j]
        elif(sub_aba_nome=='Prod-Aero'):
            temp3=itens_p[j]
        elif(sub_aba_nome=='Moto'):
            itens=itens_p
            itens[0]=str(j+1)
        else:itens=itens_p 
        if (sub_aba_nome=='Moto' or sub_aba_nome=='Coordenadas' or sub_aba_nome=='Aeros' or sub_aba_nome=='Rosas' or sub_aba_nome=='Torres'):
            wait.until(EC.element_to_be_clickable((By.XPATH, incluir)))
            wait.until(EC.element_to_be_clickable((By.XPATH, incluir))).click()
            print("incluir")
            time.sleep(15)
        else:
            if(sub_aba_nome=='Prod-Aero'):
                caminho_prod_aero=encontra_texto_no_site(num_rows,num_cols,browser,temp3[0])
                wait.until(EC.element_to_be_clickable((By.XPATH, caminho_prod_aero))).click()
                itens=temp3[1:len(temp3)]
            wait.until(EC.element_to_be_clickable((By.XPATH, editar)))
            wait.until(EC.element_to_be_clickable((By.XPATH, editar))).click()
            print("editar")
            time.sleep(15)
        if(sub_aba_nome=='Moto_s'):item_site=retorna_elementos_da_sub_aba('Moto')
        else:item_site=retorna_elementos_da_sub_aba(sub_aba_nome)
        
        if (sub_aba_nome=='Instalação'):
            if(itens!=[]): 
                if( itens[4]=='Duplo'):item_site[4]='//*[@id="ctl00_ContentPlaceHolder1_colindTipoCircuitoLTE_1"]'
        for i in range(len(itens)):
                #wait.until(EC.presence_of_element_located((By.XPATH, item_site[i])))
                wait.until(EC.visibility_of_element_located((By.XPATH, item_site[i])))
                tipo1=tipo_de_entrada(sub_aba_nome,i)
                if(tipo1=='select'):
                    time.sleep(2)
                    select_element = browser.find_element(by=By.XPATH,value=item_site[i] )#select_element=browser.find_element_by_xpath(lista[0])#subestacao
                    time.sleep(2)
                    select = Select(select_element)
                    select.select_by_visible_text(itens[i])
                    time.sleep(3.1)
                elif(tipo1=='click'):
                    wait.until(EC.presence_of_element_located((By.XPATH, item_site[i])))
                    wait.until(EC.presence_of_element_located((By.XPATH, item_site[i]))).click()
                else:
                    celula = browser.find_element(by=By.XPATH,value=item_site[i] )
                    browser.find_element(by=By.XPATH,value=item_site[i]).clear() 
                    celula.send_keys(itens[i])
                    time.sleep(2)
        time.sleep(18)
        wait.until(EC.element_to_be_clickable((By.XPATH, salvar)))
        wait.until(EC.element_to_be_clickable((By.XPATH, salvar))).click()
        print("salvar")
        time.sleep(15)
        
#######lendo quais parques o aegezitos vai ler pela planilha########## 
parques=convert_string_sem_nan(le_arquivo_excel(path, 'lista_de_processos','parques','parques'))
leilao=convert_string_sem_nan(le_arquivo_excel(path, 'lista_de_processos','parques','leilao'))
leilao_parque=[]

for j in range(len(leilao)):
     for i in range(len(parques)):
        leilao_parque.append(leilao[j]+" "+str(parques[i]))
lista_processos=[]
for i in range(len(leilao_parque)):
    lista_processos.append(convert_string_sem_nan(le_arquivo_excel(path, 'lista_de_processos','processo',leilao_parque[i])))

#print(lista_processos)

#voltar para po começo pra clicar no parque
wait = WebDriverWait(browser, 120)#definir tempo de espera caso peça login no meio do processo 

empreendimento='//*[@id="__tab_ctl00_ContentPlaceHolder1_tabEmpreendimento"]'#para clicar na aba EOL
editar='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Editar"]'
salvar='//*[@id="ctl00_ContentPlaceHolder1_ManutencaoDados1_imgButtonEPE_Salvar"]'


continuar=input('deseja continuar? (s/n): ')
#######ler a lista de parques cadastrados#########
#num_rows = len (browser.find_element(By.XPATH, "//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr"))
num_rows = len (browser.find_elements_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr"))
num_cols = len (browser.find_elements_by_xpath("//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr[2]/td"))
#num_cols = len (browser.find_element(By.XPATH,"//*[@id='ctl00_ContentPlaceHolder1_GridViewEPE1']/tbody/tr[2]/td"))
count=0
count_parques=0
count_leilao=0

##############################
continua=False
print("passo do cadastramento simplificado é 10 ")
print("passo do cadastramento completo é apartir do 1 (incluindo Outorga) ")
escolhe_item_init=input('deseja continuar apartir de um item? (t/valor numérico): ')
if(escolhe_item_init=='t'):escolhe_item_init='todos'
else:escolhe_item_init=int(escolhe_item_init)
item_leilao=0
item_parque=0
escolhe_item_parque_leilao=input('deseja continuar apartir de um parque e um leilao especifico?:(s/n) ')
if(escolhe_item_parque_leilao=='s'):
    bool_parques=input('insira o parque: ')
    for i in range(len(parques)):
        if (parques[i]==bool_parques):
            count_parques=i
    bool_leilao=input('insira o leilao: ')
    for i in range(len(leilao)):
        if (leilao[i]==bool_leilao):
            count_leilao=i
    count=(count_leilao)*len(parques)+count_parques

while continuar=='s'and count<len(lista_processos):
     start_time = time.time()
     item_leilao=leilao[count_leilao]
     item_parque=int(parques[count_parques])
     caminho=encontra_texto_no_site(num_rows,num_cols,browser,lista_processos[count][0])
     
     wait.until(EC.element_to_be_clickable((By.XPATH, caminho))).click()
     print("Parque escolhido %(n)s para o leilão %(b)s:" % {'n':item_parque ,'b':item_leilao})
     time.sleep(15)
     if(continua==False):escolhe_item=escolhe_item_init
     if (escolhe_item=='todos' or escolhe_item==1):
         aba_nome='Capacidade'
         sub_aba_nome='Capacidade'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,'entrada'))#coneferir
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
         
     if (escolhe_item=='todos' or escolhe_item==2):   
         aba_nome='Outorgas'
         sub_aba_nome='Outorgas'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,'entrada'))#coneferir
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
     if (escolhe_item=='todos' or escolhe_item==3): 
         aba_nome='Equipamentos'
         sub_aba_nome='Coordenadas'
         num_aeros=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,'Aero_parque','num aeros'))
         parque_coord=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,'Aero_parque','parque'))
         itens=[]
         if(parque_coord!=[]):
             p=0
             for i in range(len(parque_coord)):
                 if (str(item_parque)==parque_coord[i]):p=int(i)
             for i in range(1,int(num_aeros[p])+1):
                 coordenadas=str(parque_coord[p])+' '+str(i)   
                 itens.append(convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,coordenadas)))
         #print(itens)
         if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            print()
            time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
         
     if (escolhe_item=='todos' or escolhe_item==4): 
         aba_nome='Equipamentos'
         print(aba_nome)
         sub_aba_nome='Aeros'
         print(sub_aba_nome)
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,'entrada'))#coneferir
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
     
     if (escolhe_item=='todos' or escolhe_item==5): 
         aba_nome='Caracteristicas tecnicas'
         sub_aba_nome='Rosa'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1

     if (escolhe_item=='todos' or escolhe_item==6): 
         aba_nome='Caracteristicas tecnicas'
         sub_aba_nome='Local'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
         
     if (escolhe_item=='todos' or escolhe_item==7): 
         aba_nome='Caracteristicas tecnicas'
         sub_aba_nome='Torres'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
            
     if (escolhe_item=='todos' or escolhe_item==8): 
         aba_nome='Caracteristicas tecnicas'
         sub_aba_nome='Prod'
         itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
         
     if (escolhe_item=='todos' or escolhe_item==9): 
         aba_nome='Caracteristicas tecnicas'
         sub_aba_nome='Prod-Aero'
         num_aeros=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,'Aero_parque','num aeros'))
         parque_coord=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,'Aero_parque','parque'))
         itens=[]
         if(parque_coord!=[]):
             p=0
             for i in range(len(parque_coord)):
                 if (str(item_parque)==parque_coord[i]):p=int(i)
             for i in range(1,int(num_aeros[p])+1):
                 coordenadas=str(parque_coord[p])+' '+str(i)   
                 itens.append(convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,coordenadas)))
         #print(itens)
         if(itens!=[]):
             preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
             time.sleep(15)
         if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
         
     if (escolhe_item=='todos' or escolhe_item==10):
        aba_nome='Caracteristicas tecnicas'
        sub_aba_nome='Informação'
        itens=covert_string(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            #browser.save_screenshot(nome_parque+'_'+str(item_parque)+'_'+item_leilao+'.png')
            time.sleep(14)
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
        
     if (escolhe_item=='todos' or escolhe_item==11):
        aba_nome='Leilao'
        sub_aba_nome='Moto'
        item_leilao_parque=item_leilao+" "+str(item_parque)
        incluso='n'#input('Esta incluso algum item? (s/n)')
        if incluso=='s':sub_aba_nome='Moto_s'#para novos projetos , não é necessario esse if para parques novos
        itens=covert_string(le_arquivo_excel(path, aba_nome,'Moto',item_leilao_parque))
        #print(itens)
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            #browser.save_screenshot(nome_parque+'_'+str(item_parque)+'_'+item_leilao+'.png')
            time.sleep(9)
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
        
     if (escolhe_item=='todos' or escolhe_item==12):
        aba_nome='Leilao'
        sub_aba_nome='Cronograma'
        itens=covert_string(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_leilao))
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            print()
            #browser.save_screenshot(nome_parque+'_'+str(item_parque)+'_'+item_leilao+'.png')
            time.sleep(14)
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
        
     if (escolhe_item=='todos' or escolhe_item==13):
        aba_nome='Leilao'
        sub_aba_nome='Orcamento'
        itens=covert_string(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            print()
            #browser.save_screenshot(nome_parque+'_'+str(item_parque)+'_'+item_leilao+'.png')
            time.sleep(15)
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1    
        
     if (escolhe_item=='todos' or escolhe_item==14):
        aba_nome='Conexao'
        sub_aba_nome='Subestação'
        #itens=covert_string(le_arquivo_excel(path, aba_nome,sub_aba_nome,'entrada'))
        itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_leilao))
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            print()
            #browser.save_screenshot(nome_parque+'_'+str(item_parque)+'_'+item_leilao+'.png')
            time.sleep(16)
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
        
     if (escolhe_item=='todos' or escolhe_item==15):
        aba_nome='Conexao'
        sub_aba_nome='Instalação'
        itens=convert_string_sem_numeral(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_leilao))
        if(itens!=[]):
            preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
            print()
        if (escolhe_item!='todos'):escolhe_item=escolhe_item+1
     if (escolhe_item=='todos' or escolhe_item==16):
            aba_nome='Leilao'
            sub_aba_nome='REIDI'
            itens=covert_string(le_arquivo_excel(path, aba_nome,sub_aba_nome,item_parque))
            if(itens!=[]):
                preenchimento(browser,wait,aba_nome,sub_aba_nome,itens)
                print()   
     wait.until(EC.element_to_be_clickable((By.XPATH, empreendimento))).click()
     count=count+1
     count_parques=count_parques+1
     ########## conta parques e leilão #############################
     if(count_parques==len(parques)):
        count_parques=0
        count_leilao=count_leilao+1
     ################################################
     if(escolhe_item_init>9):
         escolhe_item=10
         continua=True
     time.sleep(15)   
     print("--- %s segundos de execução ---" % (time.time() - start_time))

