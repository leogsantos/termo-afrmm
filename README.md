# Chrome Mercante Automation

Este repositório contém um script Python que automatiza a interação com o site da Marinha Mercante, usando o Chrome WebDriver. Ele permite fazer o download de termos de liberação em PDF, que são normalmente usados nos portos de Ceará e Salvador.

## Configuração

Antes de executar o script, você deve configurar as informações necessárias no arquivo `data/path.txt`. Certifique-se de que ele contenha as seguintes variáveis:

- `login_mercante`: Insira o seu login de acesso à Marinha Mercante aqui.
- `senha_mercante`: Insira a sua senha de acesso à Marinha Mercante aqui.
- `url_da_marinha`: Insira a URL do site da Marinha Mercante aqui.
- `pasta_termos_pdf`: Insira o caminho da pasta onde os termos da Marinha Mercante serão salvos.

Aqui está um exemplo de como o arquivo `data/path.txt` deve ser configurado:

```plaintext
login_mercante = seu_login
senha_mercante = sua_senha
url_da_marinha = https://www.mercante.transportes.gov.br/g33159MT/jsp/logon.jsp
pasta_termos_pdf = caminho/para/pasta
```
## Uso

Certifique-se de ter o Python instalado em seu sistema. Você também deve instalar o Chrome WebDriver.
```
pip install selenium
```
```
pip install webdriver_manager
```
```
pip install pandas
```
```
pip install pdfkit
```
```
pip install winotify
```
Agora, você pode executar o script ```main.py```:

```bash
python main.py
```
O script realizará o login no site da Marinha Mercante, navegando até a seção desejada e fazendo o download de termos em PDF. Os detalhes da planilha XLSX devem ser configurados no próprio script.

## Arquivo Excel
É de comum uso dos despachantes uma referência única para cada carga que esta desembaraçando. Essa referencia deve ser inclusa na coluna A da planilha, e o número do CE mercante na coluna B

![image](https://github.com/leogsantos/termo-afrmm/assets/64739776/b9b3dc42-722e-49c9-a8a7-c6274da53956)

## Créditos

Este projeto utiliza a ferramenta de código aberto wkhtmltopdf para converter páginas da web em documentos PDF. Agradecemos à equipe do wkhtmltopdf pelo excelente trabalho na criação dessa ferramenta.

- Site do wkhtmltopdf: [https://wkhtmltopdf.org/](https://wkhtmltopdf.org/)

Certifique-se de revisar a documentação do wkhtmltopdf para obter informações detalhadas sobre seu uso, configuração e licença.
