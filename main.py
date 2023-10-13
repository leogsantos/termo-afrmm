import os
from chrome_mercante import ExecutandoChrome as Mercante

# função para obter as variaveis dentro do arquivo
def read_variables_from_file(file_path):
    variables = {}
    with open(file_path, 'r') as file:
        for line in file:
            if '=' in line:
                key, value = line.strip().split('=')
                variables[key.strip()] = value.strip()
    return variables


def main():
    # Especificar o caminho para o arquivo de texto com as variáveis
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Caminhos dinamicos
    txt_file_path = os.path.join(script_directory, 'data', 'path.txt')

    # Obter variáveis a partir do arquivo
    variables = read_variables_from_file(txt_file_path)

    # Obter variáveis para Chrome Mercante
    login_mercante = variables.get("login_mercante")
    senha_mercante = variables.get("senha_mercante")
    url_da_marinha = variables.get("url_da_marinha")
    pasta_termos_pdf = variables.get("pasta_termos_pdf")

    # recomendavel que o caminho da sua planila xlsx seja dinamico, usando a biblioteca OS.
    relatorio_afrmm_xlsx = "C:\\Insira o caminho de uma planilha com relação de referencia X Numero do Ce Mercante"

    # Inicializar e executar Chrome Mercante
    termo_marinha = Mercante(
        login_mercante, senha_mercante, url_da_marinha, pasta_termos_pdf)
    termo_marinha.main(pasta_termos_pdf, relatorio_afrmm_xlsx)

if __name__ == "__main__":
    main()