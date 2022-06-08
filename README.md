# OP-Calculadora
[![NPM](https://img.shields.io/npm/l/react)](https://github.com/LeonardoHMS/OP-Calculadora/blob/main/LICENSE)

# Sobre o projeto

OP-Calculadora é um programa com interface gráfica feito para facilitar alguns processos de trabalho com o sistema SAP. Consiste em fazer o cálculo de horas trabalhadas para o apontamento, marcando em minutos as horas trabalhas na produção e também calcular o peso de refugos.

O programa também conta com uma integração ao SAP GUI Scripting API, que auxilia a analisar as ordens de produção, buscando todas as informações necessárias pela transação, extraindo planilhas e filtrando somente as ordens que irão encerrar tecnicamente(ENTE).

# Layouts
![Window1](https://github.com/LeonardoHMS/OP-Calculadora/blob/main/assets/window.png) ![Win secondary](https://github.com/LeonardoHMS/OP-Calculadora/blob/main/assets/telas.png)

# Tecnologias utilizadas
- Python
- Json

# Bibliotecas
- PySimpleGUI
- webbrowser
- pywin32
- subprocess
- pandas
- pyperclip
- openpyxl

# Como executar o projeto
Primeiro, através do terminal, inicie o ambiente virtual do projeto:
```bash
venv\Scripts\Activate.bat
```

Depois execute o arquivo main.py:
```bash
python main.py
```

# Autor
Leonardo Henrique Mantovani Silva

www.linkedin.com/in/leonardohms
