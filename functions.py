import pandas

material = []
quantidade = []
preco = []
valorTotal = []
corte = [0.8, 0.95, 1]
individual = []
acumulada = []
classificacao = []
tabela = []

def calcularTotal(tabela):
  total = 0
  for i in range(len(tabela)):
    total += valorTotal[i]
  return total

def calcularIndividual(tabela, total):
  for i in range(len(tabela)):
    individual.append(valorTotal[i] / total)

def calcularAcumulada(tabela):
  for i in range(len(tabela)):
    if i == 0:
      acumulada.append(individual[i])
    else:
      acumulada.append(acumulada[i - 1] + individual[i])

def calcularClassificacao(tabela):
  for i in range(len(tabela)):
    if acumulada[i] <= corte[0]:
      classificacao.append('A')
    elif acumulada[i] <= corte[1]:
      classificacao.append('B')
    elif acumulada[i] <= corte[2]:
      classificacao.append('C')
    else:
      classificacao.append('C')

def imprimir(tabela):
  for i in range(len(tabela)):
    print(f"{material[i]} {quantidade[i]} {preco[i]} {valorTotal[i]:.2} {(individual[i]):.7%} {acumulada[i]:.7%} {classificacao[i]}")