Uma simples tentativa de criar uma programa que faça uma gestão de contas, deixando um grupo de contas como rotativo e outro em dia.
Este programa, lê uma planilha em excel com uma base de contas com a informação se a ação a ser tomada no sistema para aquela conta dever ser rotativo ou em dia, 
chama uma API de consulta de contas para obter as informações das contas para tomada de decisão no sistema.  E depois entra no sistema Vision Plus e realiza as transações 
necessárias para atender a ação, seja rotativo ou em dia, realizando pagamentos e débitos na conta.

Foram utilizadas as bibliotecas: 

import requests
import json
import openpyxl
import pyautogui
from time import sleep
from datetime import date
from datetime import datetime
