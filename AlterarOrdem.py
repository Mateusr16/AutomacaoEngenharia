import tkinter as tk
from tkinter import messagebox
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
import time
import datetime
import sys
import os
import glob
import csv
from webdriver_manager.chrome import ChromeDriverManager

# Configuração do Chrome
prefs = {"download.default_directory": os.getcwd()}
options = webdriver.ChromeOptions()
options.add_experimental_option("prefs", prefs)

def process_consolidado(filepath):
    """Processa o arquivo mantendo apenas as colunas 20, 22 e 23 (índices 19, 21, 22),
    inverte a primeira coluna com a segunda e retira linhas duplicadas."""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.reader(f, delimiter=';')  # Verifique o delimitador correto
            matrix = list(reader)

        if not matrix:
            print("Arquivo consolidado vazio.")
            return

        # Índices das colunas para manter (0-based)
        cols_to_keep = [19, 21, 22]
        
        # Verifica se as colunas existem
        if len(matrix[0]) < max(cols_to_keep) + 1:
            raise ValueError("O arquivo não tem colunas suficientes")

        new_matrix = []
        for row in matrix:
            try:
                # Seleciona as colunas desejadas
                new_row = [row[19], row[21], row[22]]
                # Inverte a primeira coluna com a segunda
                new_row[0], new_row[1] = new_row[1], new_row[0]
                new_matrix.append(new_row)
            except IndexError:
                print(f"Linha incompleta ignorada: {row}")
                continue

        # Remove linhas duplicadas mantendo a ordem
        unique_matrix = []
        seen = set()
        for row in new_matrix:
            row_tuple = tuple(row)
            if row_tuple not in seen:
                unique_matrix.append(row)
                seen.add(row_tuple)

        with open(filepath, 'w', newline='', encoding='utf-8') as f:
            writer = csv.writer(f, delimiter=';')  # Use o mesmo delimitador
            writer.writerows(unique_matrix)

        print("Arquivo processado. Mantidas colunas 20, 22 e 23, invertida a primeira com a segunda e duplicatas removidas.")

    except Exception as e:
        print(f"Erro ao processar consolidado: {str(e)}")


def append_data(source_path, target_path='consolidado.csv'):
    """Anexa dados usando CSV reader/writer para manter consistência"""
    try:
        # Verifica se o arquivo alvo existe
        target_exists = os.path.exists(target_path)

        with open(source_path, 'r', encoding='utf-8') as src:
            reader = csv.reader(src, delimiter=';')  # Mesmo delimitador
            rows = list(reader)

            if not rows:
                return

            header = rows[0]
            data = rows[1:]

        # Escreve usando o módulo CSV
        with open(target_path, 'a', newline='', encoding='utf-8') as tgt:
            writer = csv.writer(tgt, delimiter=';')
            
            if not target_exists:
                writer.writerow(header)
            
            writer.writerows(data)

        print(f"Dados de {source_path} anexados corretamente")

    except Exception as e:
        print(f"Erro na anexação: {str(e)}")

def wait_for_new_files(pattern, files_before, timeout=10):
    """
    Aguarda até que novos arquivos que correspondam ao padrão sejam criados,
    comparando com os arquivos já existentes (files_before).
    """
    start_time = time.time()
    while time.time() - start_time < timeout:
        files_now = set(glob.glob(pattern))
        new_files = files_now - files_before
        if new_files:
            return list(new_files)
        time.sleep(0.5)
    return []

def open_file(filepath):
    """Abre o arquivo utilizando o sistema operacional (funciona no Windows)."""
    try:
        os.startfile(filepath)
        print(f"Arquivo '{filepath}' aberto com sucesso!")
    except Exception as e:
        print(f"Erro ao abrir o arquivo {filepath}: {e}")

def exibir_palavras():
    usuario = entrada_usuario.get()
    senha = entrada_senha.get()

    if not usuario or not senha:
        messagebox.showwarning("Erro", "Por favor, preencha o usuário e a senha.")
        return

    iniciar_selenium(usuario, senha)

def iniciar_selenium(usuario, senha):
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)

    driver = webdriver.Chrome(options=options)
    driver.get('http://192.168.0.6/')

    try:
        downloaded_files = []  # Lista para armazenar os arquivos baixados

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.NAME, 'usuario')))
        print("Página de login detectada. Realizando login...")

        campo_usuario = driver.find_element(By.NAME, 'usuario')
        campo_usuario.clear()
        campo_usuario.send_keys(usuario)

        campo_senha = driver.find_element(By.NAME, 'senha')
        campo_senha.clear()
        campo_senha.send_keys(senha)
        campo_senha.send_keys(Keys.RETURN)

        time.sleep(0.1)
        print("Login efetuado com sucesso!")

        driver.get('http://192.168.0.6/pcp/plano_corte.php')
        print("Redirecionamento realizado com sucesso!")

        link_gerar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ui-id-3'))
        )
        link_gerar.click()
        print("Clicando no link 'Gerar'...")

        cod_empresa = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, 'cod_empresa'))
        )
        cod_empresa.clear()
        cod_empresa.send_keys("11")

        data_abertura = driver.find_element(By.ID, 'data_abertura')
        dia_util = datetime.date.today() - datetime.timedelta(days=1)
        while dia_util.weekday() in [5, 6]:
            dia_util -= datetime.timedelta(days=1)
        data_str = dia_util.strftime("%d-%m-%Y")
        data_abertura.clear()
        data_abertura.send_keys(data_str)
        print(f"Campo 'data_abertura' preenchido com a data de ontem: {data_str}.")

        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'cod_local'))
        )
        select_local = Select(select_element)
        select_local.select_by_value("AT MOVEIS")
        print("Opção 'AT MOVEIS' selecionado.")

        botao_gerar = driver.find_element(By.XPATH, "//input[@type='button' and @value='gerar' and @name='enviar']")
        botao_gerar.click()
        print("Botão 'gerar' clicado.")

        # Primeiro download
        files_before = set(glob.glob("plano_corte_gerado*.csv"))
        try:
            download_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='plano_corte_gerado.csv']"))
            )
            download_link.click()
            print("Link de download do arquivo CSV clicado com sucesso!")
            new_files = wait_for_new_files("plano_corte_gerado*.csv", files_before, 10)
            if new_files:
                downloaded_files.extend(new_files)
                for file in new_files:
                    open_file(file)
            else:
                print("Nenhum arquivo novo encontrado após o download.")
        except Exception as e:
            print("Download link não encontrado. Continuando o processo.")

        driver.back()
        print("Voltando à página anterior para resetar os campos.")
        WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ui-id-3'))
        )

        link_gerar = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, 'ui-id-3'))
        )
        link_gerar.click()
        print("Clicando novamente no link 'Gerar'...")

        select_element = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, 'cod_local'))
        )
        select_local = Select(select_element)
        select_local.select_by_value("PLANEJADOS")
        print("Opção 'PLANEJADOS' selecionada no campo 'cod_local'.")

        botao_gerar = driver.find_element(By.XPATH, "//input[@type='button' and @value='gerar' and @name='enviar']")
        botao_gerar.click()
        print("Botão 'gerar' clicado após selecionar 'PLANEJADOS'.")

        # Segundo download
        files_before = set(glob.glob("plano_corte_gerado*.csv"))
        try:
            download_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, "//a[@href='plano_corte_gerado.csv']"))
            )
            download_link.click()
            print("Link de download do arquivo CSV clicado com sucesso!")
            new_files = wait_for_new_files("plano_corte_gerado*.csv", files_before, 10)
            if new_files:
                downloaded_files.extend(new_files)
                for file in new_files:
                    open_file(file)
            else:
                print("Nenhum arquivo novo encontrado após o download.")
        except Exception as e:
            print("Download link não encontrado na segunda tentativa. Continuando o processo.")

        # Anexa todos os arquivos baixados ao consolidado.csv após ambos downloads
        for file in downloaded_files:
            append_data(file, 'consolidado.csv')

        # Processa o arquivo consolidado removendo colunas
        process_consolidado('consolidado.csv')

    except Exception as e:
        print(f"Erro durante o processo: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    if len(sys.argv) == 3:
        usuario = sys.argv[1]
        senha = sys.argv[2]
        iniciar_selenium(usuario, senha)
    else:
        root = tk.Tk()
        root.title("Interface de Login")
        root.geometry("400x200")

        label_usuario = tk.Label(root, text="Usuário:")
        label_usuario.pack(pady=5)
        entrada_usuario = tk.Entry(root)
        entrada_usuario.pack(pady=5)

        label_senha = tk.Label(root, text="Senha:")
        label_senha.pack(pady=5)
        entrada_senha = tk.Entry(root, show="*")
        entrada_senha.pack(pady=5)

        botao_iniciar = tk.Button(root, text="Fazer Login", command=exibir_palavras)
        botao_iniciar.pack(pady=20)

        root.mainloop()