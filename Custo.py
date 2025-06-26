import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pandas as pd

# Função para processar o Excel
def processar_excel(arquivo_excel, aba_origem="Planilha1", aba_resultado="Resultado"):

    # 1. Ler a planilha, forçando 'Data Movto' e 'Qtd Apon' como string
    df = pd.read_excel(
        arquivo_excel,
        sheet_name=aba_origem,
        converters={"Data Movto": str, "Qtd Apon": str}
    )

    # Substituir datas pela palavra "codigo" na coluna B (Data Movto)
    df["Data Movto"] = df["Data Movto"].apply(lambda x: "codigo" if pd.to_datetime(x, errors='coerce') is not pd.NaT else x)

    # 2. Converter a coluna 'Qtd Apon' para numérico
    df["Qtd Apon"] = pd.to_numeric(df["Qtd Apon"], errors="coerce")

    # Ajustar índice para facilitar a iteração
    df.reset_index(drop=True, inplace=True)

    # 3. Loop para substituir Data Movto pelo código da linha de baixo (lógica já existente)
    i = 0
    while i < len(df) - 1:
        data_movto_atual = str(df.loc[i, "Data Movto"])
        data_movto_prox = str(df.loc[i + 1, "Data Movto"])

        # Verifica se a data_movto_atual é uma data válida
        try:
            pd.to_datetime(data_movto_atual)
            is_date = True
        except ValueError:
            is_date = False

        # Se a data atual é válida e a próxima não é uma data (ou é inválida)
        if is_date and pd.isna(pd.to_datetime(data_movto_prox, errors='coerce')):
            # Substitui a Data Movto pelo código da próxima linha
            df.loc[i, "Data Movto"] = data_movto_prox
            # Remove a linha que continha o código
            df.drop(i + 1, inplace=True)
            df.reset_index(drop=True, inplace=True)
        else:
            i += 1

    # 4. Filtrar as linhas com 'Qtd Apon' >= 1
    df_filtrado = df[df["Qtd Apon"] >= 1]

    # 5. Selecionar as colunas desejadas
    colunas_desejadas = ["Documento", "Data Movto", "Qtd Apon", "Valor Realiz"]
    df_resultado = df_filtrado[colunas_desejadas]

    # 6. Salvar em outra aba
    with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_resultado.to_excel(writer, sheet_name=aba_resultado, index=False)

    # 7. Ler a planilha da aba 'Resultado' para realizar o ajuste de "codigo"
    df_resultado = pd.read_excel(arquivo_excel, sheet_name=aba_resultado)

    # Loop para substituir "codigo" pela palavra da linha de baixo na coluna B
    i = 0
    while i < len(df_resultado) - 1:
        if df_resultado.iloc[i, 1] == "codigo":
            df_resultado.iloc[i, 1] = df_resultado.iloc[i + 1, 1]
            df_resultado.drop(i + 1, inplace=True)  # Remove a linha que continha o código
            df_resultado.reset_index(drop=True, inplace=True)
        else:
            i += 1

    # Salvar novamente na aba 'Resultado'
    with pd.ExcelWriter(arquivo_excel, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df_resultado.to_excel(writer, sheet_name="Resultado", index=False)
    print("Filtragem concluída. Verifique a aba 'Resultado'.")

    return df_resultado

# Função para exibir os dados preenchidos e iniciar o script Selenium
def exibir_palavras():
    # Solicita o usuário e a senha de forma segura
    usuario = entrada_usuario.get()
    senha = entrada_senha.get()

    if not usuario or not senha:
        messagebox.showwarning("Erro", "Por favor, preencha o usuário e a senha.")
        return

    # Processa o Excel
    caminho_arquivo = 'testep.xlsx'
    df_resultado = processar_excel(caminho_arquivo)

    # Inicia o Selenium com os dados processados
    iniciar_selenium(usuario, senha, df_resultado)

# Função para iniciar o script Selenium com os dados fornecidos
def iniciar_selenium(usuario, senha, df_resultado):
    # Caminho do seu WebDriver (por exemplo, ChromeDriver)
    driver_path = r'C:\Users\Administrador\Desktop\Mateus\drive\chromedriver.exe'

    # Configurações para o ChromeDriver
    options = webdriver.ChromeOptions()

    # Criação do Service com o caminho para o ChromeDriver
    service = Service(driver_path)

    # Inicializa o navegador (aqui usando o Chrome)
    driver = webdriver.Chrome(service=service, options=options)

    # Abre a página que você quer acessar
    driver.get('http://192.168.0.6/')  # Tenta acessar a página diretamente

    # Verifica se a página de login está presente
    try:
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'usuario')))
        print("Página de login detectada. Realizando login...")

        # Preenche o campo de usuário
        campo_usuario = driver.find_element(By.NAME, 'usuario')
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        # Preenche o campo de senha
        campo_senha = driver.find_element(By.NAME, 'senha')
        campo_senha.clear()
        campo_senha.send_keys(senha)

        # Submete o formulário de login (pressiona Enter)
        campo_senha.send_keys(Keys.RETURN)

        # Aguarda um breve momento para o login ser processado
        time.sleep(1)

        # Redireciona diretamente para a página desejada após o login
        driver.get('http://192.168.0.6/pcp/man_item.php')
        print("Redirecionando para a página desejada...")

        # Preenche o campo 'cod_empresa' com o valor padrão "11"
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'cod_empresa')))
            campo_empresa = driver.find_element(By.NAME, 'cod_empresa')
            campo_empresa.clear()
            campo_empresa.send_keys("11")  # Preenche com o valor padrão
            print("Campo 'cod_empresa' preenchido com '11'.")
        except Exception as e:
            print(f"Erro ao preencher o campo 'cod_empresa': {e}")

    except Exception as e:
        print(f"Erro ao realizar login: {e}")
        driver.quit()  # Fecha o navegador em caso de erro
        return  # Encerra a função se o login falhar

    # Itera sobre todas as linhas do DataFrame
    for index, row in df_resultado.iterrows():
        print(f"Processando item {index + 1} de {len(df_resultado)}...")

        try:
            # Preenche o campo 'cod_item' com o valor da coluna 'Data Movto'
            campo_item = driver.find_element(By.NAME, 'cod_item')
            campo_item.clear()
            campo_item.send_keys(str(row['Data Movto']))
            campo_item.send_keys(Keys.RETURN)

            # Aguarda o resultado da pesquisa
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'den_item')))

            # Captura o valor do campo 'den_item'
            den_item = driver.find_element(By.NAME, 'den_item').get_attribute('value')
            print(f"Item encontrado: {den_item}")

            # Captura o valor do campo 'espf'
            espf = driver.find_element(By.NAME, 'espf').get_attribute('value')
            print(f"Especificações: {espf}")

            # Altera o campo 'custo' com o valor da coluna 'Valor Realiz'
            try:
                campo_custo = driver.find_element(By.NAME, 'custo')
                campo_custo.clear()
                campo_custo.send_keys(str(row['Valor Realiz']))
                campo_custo.send_keys(Keys.RETURN)
                # Aguarda a presença do alerta e o aceita
                WebDriverWait(driver, 10).until(EC.alert_is_present())
                alert = driver.switch_to.alert
                alert.accept()
                print(f"Campo 'custo' alterado com valor {row['Valor Realiz']} e alerta aceito.")
            except Exception as e:
                print(f"Erro ao alterar campo 'custo' para o item {index + 1}: {e}")

        except Exception as e:
            print(f"Erro ao processar item {index + 1}: {e}")

        # Volta para a página inicial para a próxima consulta
        driver.get('http://192.168.0.6/pcp/man_item.php')
        print(f"Consulta {index + 1} concluída. Voltando para a página inicial...")

        # Adiciona o preenchimento do campo 'cod_empresa' com "11" para a nova pesquisa
        try:
            WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'cod_empresa')))
            campo_empresa = driver.find_element(By.NAME, 'cod_empresa')
            campo_empresa.clear()
            campo_empresa.send_keys("11")
            print("Campo 'cod_empresa' preenchido com '11' para nova pesquisa.")
        except Exception as e:
            print(f"Erro ao preencher o campo 'cod_empresa' na nova pesquisa: {e}")

        # Aguarda um tempo antes de prosseguir para a próxima linha
        time.sleep(0.5)

    # Fecha o navegador após processar todas as linhas
    driver.quit()
    print("Processamento concluído. Navegador fechado.")

# Criando a janela principal
root = tk.Tk()
root.title("Interface de Pesquisa de Itens")

# Tamanho da janela
root.geometry("400x350")

# Criando o rótulo e campo para o usuário
label_usuario = tk.Label(root, text="Usuário:")
label_usuario.pack(pady=5)
entrada_usuario = tk.Entry(root)
entrada_usuario.pack(pady=5)

# Criando o rótulo e campo para a senha
label_senha = tk.Label(root, text="Senha:")
label_senha.pack(pady=5)
entrada_senha = tk.Entry(root, show="*")
entrada_senha.pack(pady=5)

# Botão para iniciar o script Selenium
botao_iniciar = tk.Button(root, text="Iniciar Processo", command=exibir_palavras)
botao_iniciar.pack(pady=20)

# Iniciando a interface gráfica
root.mainloop()
