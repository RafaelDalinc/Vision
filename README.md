Uma simples tentativa de criar um programa para controlar uma massa de dados no sistema Vision Plus.

O programa lê uma planilha em excel contendo os cartões e campo de ação: rotativou ou em dia.  
Após ler as contas, chama uma API de consulta para obter as informações necessárias para tomar decisão se realizar pagamento ou compras para os cartões.

Linguagem de programação: Python

Bibliotecas usadas:
import requests
import json
import openpyxl
import pyautogui
from time import sleep
from datetime import date
from datetime import datetime
