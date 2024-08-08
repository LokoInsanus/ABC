import os
import openpyxl
import openpyxl.styles as styles
import matplotlib.pyplot as plt
from matplotlib.ticker import FormatStrFormatter

def calcularTotal(valorTotal):
  total = 0
  for i in range(len(valorTotal)):
    total += valorTotal[i]
  return total

def calcularIndividual(valorTotal, total):
  individual = []
  for i in range(len(valorTotal)):
    individual.append(valorTotal[i] / total)
  return individual

def calcularaAcumulada(individual):
  acumulada = []
  for i in range(len(individual)):
    if i == 0:
      acumulada.append(individual[i])
    else:
      acumulada.append(acumulada[i - 1] + individual[i])
  return acumulada

def calcularClassificacao(corte: list, acumulada:list):
  classificacao = []
  for i in range(len(acumulada)):
    if acumulada[i] <= corte[0]:
      classificacao.append('A')
    elif acumulada[i] <= corte[1]:
      classificacao.append('B')
    elif acumulada[i] <= corte[2]:
      classificacao.append('C')
    else:
      classificacao.append('C')
  return classificacao

def propSKU(classificacao):
  propA = 0
  propB = 0
  propC = 0
  tamanho = len(classificacao)
  for i in range(tamanho):
    if classificacao[i] == 'A':
      propA += 1
    elif classificacao[i] == 'B':
      propB += 1
    elif classificacao[i] == 'C':
      propC += 1
  propA /= tamanho
  propB /= tamanho
  propC /= tamanho
  return propA, propB, propC

def propValor(classificacao, individual):
  propA = 0
  propB = 0
  propC = 0
  for i in range(len(classificacao)):
    if classificacao[i] == 'A':
      propA += individual[i]
    elif classificacao[i] == 'B':
      propB += individual[i]
    elif classificacao[i] == 'C':
      propC += individual[i]
  return propA, propB, propC

def gerarGraficos(individual, acumulada):
  fig, grafico = plt.subplots()
  op = int(input("1 - Individual\n2 - Acumulada\nEscolha uma das opções acimea: "))
  #grafico.yaxis.set_major_formatter(FormatStrFormatter('%d'))
  match op:
    case 1:
      individualPorcentagem = []
      listaTotal = [i for i in range(0, len(individual))]
      for i in individual:
        individualPorcentagem.append(i * 100)
      grafico.bar(listaTotal, individualPorcentagem)
      media = acharMedia(individualPorcentagem)
      grafico.set_ylim(0, media)
    case 2:
      acumuladaPorcentagem = []
      for i in acumulada:
        acumuladaPorcentagem.append(i * 100)
      grafico.plot(acumuladaPorcentagem)
    case _:
      print("Opção invalida")
  #fig.tight_layout()
  plt.show()

def gerarTabela(individual, acumulada, classificacao, corte, material, quantidade, preco, valorTotal):
  abc = ['A', 'B', 'C']
  tamanho = len(classificacao)
  skus = propSKU(classificacao)
  valor = propValor(classificacao, individual)
  listaCortes = [f'{corte[0]:.2%}', f'{corte[1]:.2%}', f'{corte[2]:.2%}']
  listaSKUs = [f'{skus[0]:.2%}', f'{skus[1]:.2%}', f'{skus[2]:.2%}']
  listaValores = [f'{valor[0]:.2%}', f'{valor[1]:.2%}', f'{valor[2]:.2%}']
  for i in range(len(individual)):
    individual[i] = f"{individual[i]:.2%}"
  for i in range(len(acumulada)):
    acumulada[i] = f"{acumulada[i]:.2%}"
  listaCortes += [''] * (tamanho - 3)
  listaSKUs += [''] * (tamanho - 3)
  listaValores += [''] * (tamanho - 3)
  abc += [''] * (tamanho - 3)
  dados = {'Material': material,
          'Quantidade': quantidade,
          'Preço': preco,
          'Valor Total': valorTotal,
          'Individual': individual,
          'Acumulada': acumulada,
          'Classificação': classificacao,
          '': '',
          'Classificaçao': abc,
          'Corte': listaCortes,
          'Proporção de SKUs': listaSKUs,
          'Proporção de Valor': listaValores
          }
  return dados

def configurarTabela(nome, tamanho):
  planilha = openpyxl.load_workbook(f'{nome}.xlsx')
  tabela = planilha['Sheet1']
  tabela.column_dimensions['A'].width = 50
  tabela.column_dimensions['B'].width = 15
  tabela.column_dimensions['C'].width = 15
  tabela.column_dimensions['D'].width = 15
  tabela.column_dimensions['E'].width = 15
  tabela.column_dimensions['F'].width = 15
  tabela.column_dimensions['G'].width = 15
  tabela.column_dimensions['I'].width = 15
  tabela.column_dimensions['J'].width = 10
  tabela.column_dimensions['K'].width = 20
  tabela.column_dimensions['L'].width = 20
  tabela['H1'].border = styles.Border(left=styles.Side(border_style='thin'), right=styles.Side(border_style='thin'))
  #tabela[f'C2:D{tamanho}'].number_format = styles.NumberFormat('$#,##0.00')
  alinhamento = styles.Alignment(horizontal='center', vertical='center')
  intervalo = tabela[f'B2:G{tamanho + 1}']
  for linha in intervalo:
    for célula in linha:
      célula.alignment = alinhamento
  intervalo = tabela['I2:L5']
  for linha in intervalo:
    for célula in linha:
      célula.alignment = alinhamento
  planilha.save(f'{nome}.xlsx')
  planilha.close()

def apagarTela():
  if os.name == "posix":
    os.system('clear')
  elif os.name == "nt":
    os.system('cls')

def acharMedia(lista):
  media = sum(lista)
  return media / len(lista)