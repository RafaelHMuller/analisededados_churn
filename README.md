<h1 align="center">
📄<br>README - Projeto de Análise de Dados
</h1>

## Índice 

* [Descrição do Projeto](#descrição-do-projeto)
* [Funcionalidades e Demonstração da Aplicação](#funcionalidades-e-demonstração-da-aplicação)
* [Pré requisitos](#pré-requisitos)
* [Execução](#execução)
* [Bibliotecas](#bibliotecas)

# Descrição do projeto
> Este é o repositório do meu projeto de análise de dados. Meu objetivo foi treinar atutomações online e no computador, análise de dados e envio automático de e-mail. Inicialmente, por meio de automações, buscou-se uma base de dados online de uma empresa de telecomunicações fictícia. Então, após a coleta da base de dados, via uma análise de dados, procurou entender os principais motivos dos cancelamentos dos usuários (churn). Por fim, enviou-se um email à diretoria da empresa com gráficos e arquivo de texto contendo as explicações dos principais motivos encontrados para o número de churn.

# Funcionalidades e Demonstração da Aplicação
Envio de email com informações dos principais motivos que levaram aos cancelamentos dos clientes:
- arquivo de texto Word com toda a análise de dados
- gráficos anexados com a comparação entre as variáveis/features da base de dados

![Screenshot_3](https://user-images.githubusercontent.com/128300382/227017029-ced9cd3e-8103-41b4-8287-3ac068f85201.png)

Imagem do arquivo de texto Word com as informações coletadas a partir da análise de dados:

![Screenshot_4](https://user-images.githubusercontent.com/128300382/227017052-d2eac2c6-96d6-4d84-ac6a-90ddcd82c88f.png)


## Pré requisitos

* Sistema operacional Windows
* IDE de python (ambiente de desenvolvimento integrado de python)
* Navegador Google Chrome

## Execução

Neste projeto, há automação (selenium e pyautogui). Durante a automação web, uma mensagem de alerta será mostrada ao usuário, recomendando a não utilização do teclado ou mouse. Ao fim desta automação, outra mensagem de alerta indicará seu final.

## Bibliotecas

* selenium: biblioteca de automação web
* webdriver_manager.chrome: em conjunto com o selenium, atualiza o drive do Chrome
* pyautogui: biblioteca de automação por meio do mouse, teclado e monitor
* pandas: biblioteca de análise de dados
* win32com.client: biblioteca que permite a utilização de aplicações do Windows (ex.: Outlook)
* time: biblioteca que permite definir intervalos de pausa na automação
* os: biblioteca de integração de arquivos e pastas do computador
* shutil: biblioteca utilizada para mover arquivos através de pastas
* plotly.express: biblioteca de criação de gráficos
* docx: biblioteca que permite a integração a arquivos Word
* datetime: biblioteca que permite a utilização de datas e horários
