import pandas
import matplotlib.pyplot as plt
from tabulate import tabulate

caminho = './Estoque Almoxarifado.xlsx'
tabela = pandas.read_excel(caminho)
material = tabela['Material'].tolist()
quantidade = tabela['Quantidade'].tolist()
preco = tabela['Preço'].tolist()
valorTotal = tabela['Valor Total'].tolist()
total = 0
op = 0
corte = [0.8, 0.95, 1]
individual = []
acumulada = []
classificacao = []

def calcularClassificacao():
    for i in range(len(tabela)):
        if acumulada[i] <= corte[0]:
            classificacao.append('A')
        elif acumulada[i] <= corte[1]:
            classificacao.append('B')
        elif acumulada[i] <= corte[2]:
            classificacao.append('C')
        else:
            classificacao.append('C')

# caminho = input("Digite o nome do arquivo: ")
# caminho += ".xlsx"

for i in range(len(tabela)):
    total += valorTotal[i]
for i in range(len(tabela)):
    individual.append(valorTotal[i] / total)
for i in range(len(tabela)):
    if i == 0:
        acumulada.append(individual[i])
    else:
        acumulada.append(acumulada[i - 1] + individual[i])
calcularClassificacao()

while op != 6:
    op = int(input("1 - Modificar Corte\n2 - Proporção de SKUs\n3 - Proporção de Valor\n4 - Imprimir Tabela\n5 - Gerar Grafico\n6 - Sair\nEscolha uma das opções acima: "))
    match op:
        case 1:
            classificacao = []
            corte[0] = float(input("Corte A: ")) / 100
            corte[1] = float(input("Corte B: ")) / 100
            corte[2] = float(input("Corte C: ")) / 100
            calcularClassificacao()
        case 2:
            propA = 0
            propB = 0
            propC = 0
            total = len(tabela)
            for i in range(total):
                if classificacao[i] == 'A':
                    propA += 1
                elif classificacao[i] == 'B':
                    propB += 1
                elif classificacao[i] == 'C':
                    propC += 1
            print(f"{(propA / total):.2%}")
            print(f"{(propB / total):.2%}")
            print(f"{(propC / total):.2%}")
        case 3:
            propA = 0
            propB = 0
            propC = 0
            for i in range(len(tabela)):
                if classificacao[i] == 'A':
                    propA += individual[i]
                elif classificacao[i] == 'B':
                    propB += individual[i]
                elif classificacao[i] == 'C':
                    propC += individual[i]
            print(f"{propA:.2%}")
            print(f"{propB:.2%}")
            print(f"{propC:.2%}")
        case 4:
            for i in range(10):
                print(f"{material[i]} {quantidade[i]} {preco[i]} {valorTotal[i]:.2} {(individual[i]):.7%} {acumulada[i]:.7%} {classificacao[i]}")
        case 5:
            plt.bar(len(material), acumulada)
            plt.show()
        case _:
            if op != 6:
                print("Opção Invalida")