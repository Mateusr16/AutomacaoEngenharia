# app.py
from flask import Flask, render_template, request, redirect, url_for, jsonify
import subprocess
import os
import sys
import json
import tempfile
import pandas as pd
import time

def resource_path(rel_path):
    """ Get absolute path to resource for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, rel_path)

app = Flask(__name__, template_folder=resource_path('templates'))

# Rota para a página inicial
@app.route('/')
def index():
    return render_template('main.html')

# Rota para mostrar o formulário de alterar descrição
@app.route('/formulario_alterar_descricao', methods=['GET'])
def formulario_alterar_descricao():
    return render_template('alterar_descricao.html')

# Rota para executar o script alterardescricao.py
@app.route('/executar_alterar_descricao', methods=['POST'])
def executar_alterar_descricao():
    usuario = request.form['usuario']
    senha = request.form['senha']
    campos_selecionados = ','.join(request.form.getlist('campos'))
    dados_tabela = request.form.get('dados_tabela')
    
    if not campos_selecionados:
        return "Selecione pelo menos um campo!", 400
    
    # Criar arquivo temporário com os dados da tabela
    try:
        dados = json.loads(dados_tabela)
        df = pd.DataFrame(dados)
        
        # Garantir colunas obrigatórias
        if 'cod_empresa' not in df.columns:
            df['cod_empresa'] = ''
        if 'cod_item' not in df.columns:
            df['cod_item'] = ''
        
        # Criar arquivo temporário
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        temp_path = temp_file.name
        df.to_excel(temp_path, index=False)
        temp_file.close()
    except Exception as e:
        return f"Erro ao processar dados: {str(e)}", 500

    script_path = resource_path('alterardescricao.py')
    if os.path.exists(script_path):
        try:
            python_exec = sys.executable
            
            # Executar e capturar a saída
            process = subprocess.Popen(
                [python_exec, script_path, usuario, senha, campos_selecionados, temp_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True
            )
            
            # Aguardar término do processo
            stdout, stderr = process.communicate()
            
            # Tentar extrair o relatório JSON
            relatorio = {}
            try:
                # Procurar a última linha que contém JSON
                for line in stdout.splitlines():
                    if line.startswith('{'):
                        relatorio = json.loads(line)
            except json.JSONDecodeError:
                relatorio = {
                    "error": "Erro ao analisar relatório",
                    "stdout": stdout,
                    "stderr": stderr
                }
            
            # Adicionar caminho do relatório ao resultado
            relatorio["relatorio_path"] = temp_path.replace('.xlsx', '_relatorio.json')
            
            return jsonify(relatorio)
            
        except Exception as e:
            # Remover arquivo temporário em caso de erro
            if os.path.exists(temp_path):
                os.unlink(temp_path)
            return jsonify({
                "error": f"Erro ao iniciar o script: {e}",
                "traceback": str(e)
            }), 500
    return jsonify({"error": "Script não encontrado"}), 404

# Rota para executar o script AlterarOrdem.py
@app.route('/executar_alterar_ordem', methods=['POST'])
def executar_alterar_ordem():
    script_path = resource_path('AlterarOrdem.py')
    executar_script(script_path)
    return redirect(url_for('index'))

# Rota para executar o script Custo.py
@app.route('/executar_custo', methods=['POST'])
def executar_custo():
    script_path = resource_path('Custo.py')
    executar_script(script_path)
    return redirect(url_for('index'))

# Função auxiliar para executar scripts
def executar_script(script_path):
    if os.path.exists(script_path):
        try:
            python_exec = sys.executable
            subprocess.Popen([python_exec, script_path])
            print(f"Script '{script_path}' iniciado com sucesso!")
        except Exception as e:
            print(f"Erro ao iniciar o script: {e}")
    else:
        print(f"Erro: Script não encontrado em {script_path}")

# Adicione isso ao final do app.py para uso com Waitress
from waitress import serve

if __name__ == '__main__':
    serve(app, host='0.0.0.0', port=5000)