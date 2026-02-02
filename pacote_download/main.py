from selenium.webdriver import Chrome
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import NoSuchElementException
from selenium.webdriver.support.ui import Select
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from time import sleep
from datetime import datetime
import os
import pandas as pd
from dotenv import load_dotenv
load_dotenv()

# ======= CONFIGURAÇÕES INICIAIS =======
def configuracoes_chrome():
    chrome_options = Options()
    chrome_options.add_argument("--log-level=3")
    chrome_options.add_experimental_option("excludeSwitches", ["enable-logging"])
    chrome_options.add_argument("--start-maximized")
    return chrome_options


# ======= DATA =======
hoje = datetime.today()
mes_atual = hoje.month
ano_atual = hoje.year

# ===== VERIFICAR MÊS PASSADO =====
if mes_atual == 1:
    mes_passado = 12
    ano_passado = ano_atual - 1
else:
    mes_passado = mes_atual - 1
    ano_passado = ano_atual

config = {
    "fechamento_25": f"25/{mes_atual:02d}/{ano_atual}",
    "fechamento_10": f"10/{mes_atual:02d}/{ano_atual}",
    "usuario": os.getenv("USUARIO"),
    "senha": os.getenv("SENHA"),
    "arquivo_excel": r"C:\Users\DELL\Documents\Solicitações\Controladoria\Amanda\analise agregados\dados_agregados\dados_agregados.xlsx",
    "saida_excel": r"C:\Users\DELL\Documents\Solicitações\Controladoria\Amanda\analise agregados\Dezembro\1 QUINZENA DE JANEIRO.xlsx",
    "natureza": ["11"],
    "filiais_excluidas": ["SAO PAULO", "RECIFE"],
    "url_brudam": "https://vdclog.brudam.com.br/financeiro/contas_pagar.php?"
}

meses = {
    1 : "JANEIRO", 2 : "FEVEREIRO", 3 : "MARÇO", 4 : "ABRIL", 5 : "MAIO", 6 : "JUNHO", 7 : "JULHO", 8 : "AGOSTO", 9 : "SETEMBRO", 10 : "OUTUBRO", 11 : "NOVEMBRO", 12 : "DEZEMBRO"
}


# ======= ENTRAR NO SISTEMA =======
chrome_options = configuracoes_chrome()
service = Service(log_path="NUL")
browser = webdriver.Chrome(service=service, options=chrome_options)
url = config["url_brudam"]
browser.get(url)

# ======= REGISTRO DE ESPERA =======
wdw = WebDriverWait(browser, 10)

# ======= DATAFRAME COM DADOS DOS AGREGADOS =======
df = pd.read_excel(config["arquivo_excel"])
df["CUSTO"] = df["CUSTO"].fillna("").astype(str)
df["MANIF."] = df["MANIF."].fillna("").astype(str)

# ======= VERIFICAR DIA =======
if hoje.day > 14:
    fechamento = 25
    data_inicial = data_final =  config["fechamento_25"]
else:
    fechamento = 10
    data_inicial = data_final =  config["fechamento_10"]


# ======= VERIFICAR DATA DE FECHAMENTO =======
if fechamento == 10:
    quinzena_atual = f"2 QUINZENA DE {meses[mes_passado]}"
    quinzena_anterior = f"1 QUINZENA DE {meses[mes_passado]}"
    mensal_anterior = f"MENSAL - {meses[mes_passado]}"
else:
    quinzena_atual = f"1 QUINZENA DE {meses[mes_atual]}"
    quinzena_anterior = f"2 QUINZENA DE {meses[mes_passado]}"
    mensal_anterior = ""

# ======= FUNÇÃO PARA CONVERTER VALOR =======
def converter_valor(valor):
    return float(valor.replace('.', '').replace(',','.'))

# ======= FUNÇÃO DE LOGIN DO USUÁRIO =======
def login(usuario, senha):
    browser.find_element(By.ID, "user").send_keys(usuario)
    browser.find_element(By.ID, "password").send_keys(senha)
    browser.find_element(By.ID, "acessar").click()


# ======= FUNÇÃO PARA ENCONTRAR ELEMENTO =======
def elemento(encontrar, elemento, valor=None):
    # ======= ENCONTRAR ELEMENTO, CLICAR =======
    if valor == None:
        if encontrar == "ID":
            elemento_atual = browser.find_element(By.ID, elemento)
            elemento_atual.click()

        elif encontrar == "CSS":
            elemento_atual = browser.find_element(By.CSS_SELECTOR, elemento)
            elemento_atual.click()

        elif encontrar == "XPATH":
            elemento_atual = browser.find_element(By.XPATH, elemento)
            elemento_atual.click()

    # ======= ENCONTRAR ELEMENTO, ENVIAR VALOR  =======
    else:
        if encontrar == "ID":
            elemento_atual = browser.find_element(By.ID, elemento)
            elemento_atual.click()
            elemento_atual.send_keys(valor)

        elif encontrar == "CSS":
            elemento_atual = browser.find_element(By.CSS_SELECTOR, elemento)
            elemento_atual.click()
            elemento_atual.send_keys(valor)


# ======= FUNÇÃO PARA VERIFICAR DATA LIMITE =======
def existe_data_maior(manifesto_saida, data_limite_str):
    data_limite = datetime.strptime(data_limite_str, "%d/%m/%Y")
    for data in manifesto_saida:
        if not data.strip():
            continue
        try:
            data_convertida = datetime.strptime(data, "%d/%m/%Y")
        except ValueError:
            continue
        if data_convertida > data_limite:
            return True
    return False


# ======= REALIZAR LOGIN =======
wdw.until(EC.element_to_be_clickable((By.ID, "user")))
login(config["usuario"], config["senha"])

# ======= ESPERAR FORM1 ESTAR PRESENTE =======
wdw.until(EC.presence_of_element_located((By.ID, "form1")))

# ======= CONFIGURAR DATA =======
wdw.until(EC.element_to_be_clickable((By.ID, "data_1")))
elemento("ID", "data_1", data_inicial)
elemento("ID", "data_2", data_final)

# ======= SELECIONAR SITUAÇÃO DA FATURA =======
wdw.until(EC.presence_of_element_located((By.ID,"situacao")))
elemento("ID", "selSituacao")
elemento("CSS", "input[type='checkbox'][value='5']")
elemento("CSS", ".ui-dialog-buttonset")

# ======= ESPERAR CENTRO DE CUSTO APARECER =======
elemento("ID", "selecionaCentros")
wdw.until(EC.presence_of_element_located((By.ID, "formularioNovoCentroCusto")))
# ======= SELECIONAR CENTRO DE CUSTO =======
for item in config["natureza"]:
    centro_custo = browser.find_element(By.CSS_SELECTOR, f"input[type='checkbox'][class='check_centro'][value='{item}']")
    centro_custo.click()
elemento("XPATH", "//button[text()='Incluir']")

# ======= FILTRAR FILIAL =======
elemento("ID", "filtro_unidades")
wdw.until(
    EC.presence_of_element_located(
        (By.ID, "formularioUnidades")
    )
)
for item in config["filiais_excluidas"]:
    unidades = browser.find_elements(By.CSS_SELECTOR, f"input[textcheck='{item}']")
    for unidade in unidades:
        unidade.click()
elemento("XPATH", "//button[text()='Incluir']")

# ======= PESQUISAR FATURAS =======
elemento("ID", "PESQUISAR")

# ======= ESPERAR OS LANÇAMENTOS APARECEREM =======
wdw.until(EC.presence_of_element_located((By.ID, "listaLancamentos")))
LinhasTabela = browser.find_elements(By.CLASS_NAME, "LinhaTabela")
aba_principal = browser.current_window_handle

# ======= VERIFICAR CÓDIGOS VISITADOS PARA NÃO REPETI-LOS =======
adicionar_codigos = []
with open("arquivo.txt", "r", encoding="utf-8") as codigos_visitados:
    codigos_visitados = {linha.strip() for linha in codigos_visitados}


# ======= ENTRAR NOS LANÇAMENTOS PELO CÓDIGO =======
for linha in LinhasTabela:
    erro_quinzena = ""
    situacao_manifesto = ""
    manifesto_repetido = ""
    lista_datas_saida = []
    manifesto_saida = []
    manifesto_placa = []
    manifesto_finalizado = []
    manifesto_valor = []
    valor_agregado = []
    vlr_rota = ""
    status_manifesto = "OK"
    status_custo = ""
    
    # ======= ENTRAR NOS LANÇAMENTOS =======
    id_lancamento = linha.get_attribute("id")
    if id_lancamento in codigos_visitados:
        continue
    url = f"https://vdclog.brudam.com.br/financeiro/consulta_lancamento.php?id={id_lancamento}"
    browser.execute_script(f"window.open(arguments[0]);",url)
    browser.switch_to.window(browser.window_handles[-1])

    # ======= VALOR DA FATURA =======
    wdw.until(EC.presence_of_element_located((By.ID,"naturezaValor_11")))
    valor_fatura = float(browser.find_element(By.ID, "naturezaValor_11").get_attribute("value"))

    # ======= VERIFICAR SE LANÇAMENTO JÁ FOI ANALISADO =======
    adicionar_codigos.append(id_lancamento)
    wdw.until(EC.presence_of_element_located((By.ID,"formularioLancamento")))
    sleep(3)
    
    # ======= PEGAR NOME AGREGADO =======
    fornecedor_nome = browser.find_element(By.ID, "fornecedor_bold")
    fornecedor_nome = fornecedor_nome.get_attribute("value")
    acordo_agregado = df.loc[df["AGREGADO"] == fornecedor_nome, "ACORDO"]
    acordo_agregado = list(acordo_agregado)


    # ======= ABRIR FATURA =======
    try:
        elemento("ID","abrirFatura")
        browser.switch_to.window(browser.window_handles[-1])
        # Listar manifestos
        wdw.until(EC.presence_of_element_located((By.CLASS_NAME, "selecao")))
        selecoes = browser.find_elements(By.CLASS_NAME, "selecao")
        for selecao in selecoes:
            lista_linhas = selecao.find_elements(By.CLASS_NAME, "LISTA_linha")
            manifesto_placa.append(lista_linhas[2].text)
            manifesto_saida.append(lista_linhas[3].text)
            manifesto_finalizado.append(lista_linhas[4].text)
            manifesto_valor.append(lista_linhas[11].text)

        # ======= VERIFICAR SITUAÇÃO DOS MANIFESTOS =======
        if "" in manifesto_placa:
            situacao_manifesto += "S/PLACA - "
            status_manifesto = "NOK"
        if "" in manifesto_saida:
            situacao_manifesto += "MANIFESTO(S) S/SAIDA - "
            status_manifesto = "NOK"
        if "" in manifesto_finalizado:
            situacao_manifesto += "MANIFESTO(S) S/FECHAR - "
            status_manifesto = "NOK"
        else:
            status_manifesto = "OK"

              # ======= VERIFICAR SE AGREGADO FOI ENCONTRADO =======
        if acordo_agregado == []:
            pass
        else:
            # ======= VERIFICAR QUINZENA / MENSAL =======
            if "QUINZENA" in acordo_agregado or "MENSAL" in acordo_agregado:
                valor_agregado = df.loc[df["AGREGADO"] == fornecedor_nome, "VALOR"].unique()
                valor_agregado = converter_valor(str(valor_agregado[0]))
                if valor_fatura == valor_agregado:
                    vlr_rota = ""
                    status_custo = "OK"
                else:
                    vlr_rota = "VLR ROTA - "
                    status_custo = "NOK"
            # ======= VERIFICAR DIÁRIA =======
            elif "DIÁRIA" in acordo_agregado:
                valor_agregado = df.loc[df["AGREGADO"] == fornecedor_nome, "VALOR"].unique()
                valor_agregado = converter_valor(str(valor_agregado[0]))
                valores_manifesto = [converter_valor(valor) for valor in manifesto_valor]

                for data in manifesto_saida:
                    data_convertida = datetime.strptime(data, "%d/%m/%Y")
                    lista_datas_saida.append(data_convertida)
                lista_datas_unicas = set(lista_datas_saida)
                if len(lista_datas_saida) != len(lista_datas_unicas):
                    manifesto_repetido = ("DIA_MAN - ")
                if any(valor != valor_agregado for valor in valores_manifesto):
                    vlr_rota = "VLR ROTA - "
                    status_custo = "NOK"
                else:
                    status_custo = "OK"
                    vlr_rota = ""
            else:
                vlr_rota = "ANALISE"

        # ======= FECHAR ABA E RETORNAR PARA A ABA DO LANÇAMENTO =======
        browser.close()
        browser.switch_to.window(browser.window_handles[-1])
    except:
        print(f"Código {id_lancamento} não tem fatura!")

    # ======= VERIFICAR SE A PÁGINA TEM ALGUM ANEXO (CONSIDERANDO COMO NF) =======
    try:
        browser.find_element(By.CSS_SELECTOR, "i.fa.fa-download.fa-xs")
        existe_download = True
    except NoSuchElementException:
        existe_download = False
    verificar_nf = "NF+" if existe_download else "SEM NF"

    # ======= VERIFICAR COMPETÊNCIA CADASTRADA AO AGREGADO =======
    mapa_competencia = {
        "QUINZENA ATUAL": quinzena_atual,
        "QUINZENA ANTERIOR": quinzena_anterior,
        "MENSAL - ANTERIOR": mensal_anterior
        }
    resultado = df[df['AGREGADO'] == fornecedor_nome]

    # ======= ACORDO DO AGREGADO NÃO ESTÁ LISTADO =======
    if vlr_rota == "ANALISE":
        wdw.until(EC.presence_of_element_located((By.ID,"descricao")))
        descricao = browser.find_element(By.ID, "descricao")
        descricao.clear()
        descricao.send_keys("ANALISE ACORDO")

    # ======= SE O NOME DO AGREGADO ESTÁ IGUAL AO SISTEMA FAZ ISSO =======
    elif not resultado.empty:
        competencia = resultado.iloc[0]['COMPETÊNCIA'].strip()
        descricao_texto = mapa_competencia.get(competencia)

        # ======= VERIFICAR QUINZENA =======
        if "1 QUINZENA" in descricao_texto:
            if fechamento == 10:
                data_limite = f"15/{mes_passado:02d}/{ano_passado}"
            else:
                data_limite = f"15/{mes_atual:02d}/{ano_atual}"
            if existe_data_maior(manifesto_saida, data_limite):
                erro_quinzena = " ERRO QUINZ. -"

        elif "2 QUINZENA" in descricao_texto:
            data_limite = f"31/{mes_passado:02d}/{ano_passado}"
            if existe_data_maior(manifesto_saida, data_limite):
                erro_quinzena = " ERRO QUINZ. -"

        # ======= VERIFICAR MÊS COMPETÊNCIA =======
        wdw.until(EC.presence_of_element_located((By.ID, "competencia")))
        if fechamento == 10:
            browser.execute_script(
                f"document.getElementById('competencia').value = '{ano_passado}-{mes_passado:02d}';"
            )
        else:
            if competencia == "QUINZENA ANTERIOR":
                browser.execute_script(
                    f"document.getElementById('competencia').value = '{ano_passado}-{mes_passado:02d}';"
                )
            elif competencia == "QUINZENA ATUAL":
                browser.execute_script(
                    f"document.getElementById('competencia').value = '{ano_passado}-{mes_atual:02d}';"
                )
            else:
                browser.execute_script(
                    f"document.getElementById('competencia').value = '{ano_passado}-{mes_passado:02d}';"
                )

        # ======= ESCREVER NA DESCRIÇÃO DA FATURA =======
        wdw.until(EC.presence_of_element_located((By.ID,"descricao")))
        descricao = browser.find_element(By.ID, "descricao")
        descricao.clear()
        descricao.send_keys(f"{descricao_texto} - TORRE DE CONTROLE - {manifesto_repetido}{vlr_rota}{verificar_nf} -{erro_quinzena} {situacao_manifesto}PAGAMENTO A TERCEIROS TAC: X")

        df.loc[df["AGREGADO"] == fornecedor_nome, "CUSTO"] = status_custo
        df.loc[df["AGREGADO"] == fornecedor_nome, "MANIF."] = status_manifesto
        df.loc[df["AGREGADO"] == fornecedor_nome, "VLR FATURA"] = valor_fatura

        df_novo = df.copy()
        df_novo.to_excel(config["saida_excel"], index=False)
        print("Arquivo Excel Gerado com Sucesso!")
        
    # ======= SE NAO ENCONTROU AGREGADO OU É DUPLICADO PÕEM COMO ANÁLISE =======
    else:
        wdw.until(EC.presence_of_element_located((By.ID,"descricao")))
        descricao = browser.find_element(By.ID, "descricao")
        descricao.clear()
        descricao.send_keys("ANALISE NOME")

    # ======= VERIFICAR FORMA DE PAGAMENTO =======
    wdw.until(EC.presence_of_element_located((By.ID, "forma")))
    forma_pagamento = browser.find_element(By.ID, "forma")
    selecionar_forma_pagamento = Select(forma_pagamento)
    selecionar_forma_pagamento.select_by_value("13")

    # ======= SELECIONAR CONTA BANCÁRIA =======
    wdw.until(EC.presence_of_element_located((By.ID, "conta_bancaria_contra")))
    conta_bancaria_contra = browser.find_element(By.ID, "conta_bancaria_contra")
    conta_bancaria = []
    selecionar_conta_bancaria = Select(conta_bancaria_contra)

    for opcao in selecionar_conta_bancaria.options:
        numero_conta_bancaria = opcao.get_attribute("value")
        conta_bancaria.append(numero_conta_bancaria)
    
    if len(conta_bancaria) > 1:
        conta_bancaria.sort()
        selecionar_conta_bancaria.select_by_value(conta_bancaria[1])

    # ======= SALVAR PROCESSO DA FATURA =======
    salvar = browser.find_element(By.ID, "salvar").click()
    sleep(2)
    wdw.until(EC.invisibility_of_element_located((By.ID,"showMessageBrudam")))

    # ======= FECHAR ABA E RETORNAR A ABA DE LANÇAMENTOS =======
    browser.close()
    browser.switch_to.window(aba_principal)
    
# ======= SAIR DO PROGRAMA =======
browser.quit()

# ======= SALVAR CÓDIGOS JÁ VISITADOS =======
with open("arquivo.txt", "a", encoding="utf-8") as codigos_visitados:
    codigos_visitados.writelines(f"{codigo}\n" for codigo in adicionar_codigos)