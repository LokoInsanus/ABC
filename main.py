import pandas
import matplotlib.pyplot as plt
from tabulate import tabulate
import subprocess

#caminho = './Estoque Almoxarifado.xlsx'
caminho = input("Digite o nome do arquivo: ")
caminho += ".xlsx"
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

def propSKU():
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
    propA /= total
    propB /= total
    propC /= total
    return propA, propB, propC

def propValor():
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
    return propA, propB, propC

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

while op != 7:
    op = int(input("1 - Modificar Corte\n2 - Proporção de SKUs\n3 - Proporção de Valor\n4 - Imprimir Tabela\n5 - Gerar Grafico\n6 - Exportar Tabela\n7 - Sair\nEscolha uma das opções acima: "))
    match op:
        case 1:
            classificacao = []
            corte[0] = float(input("Corte A: ")) / 100
            corte[1] = float(input("Corte B: ")) / 100
            corte[2] = float(input("Corte C: ")) / 100
            calcularClassificacao()
        case 2:
            a, b, c = propSKU()
            print(f"A: {a:.2%}\nB: {b:.2%}\nC: {c:.2%}")
        case 3:
            a, b, c = propValor()
            print(f"A: {a:.2%}\nB: {b:.2%}\nC: {c:.2%}")
        case 4:
            print("Digite um intevalo para mostra os dados, sendo no máximo 1000")
            primeiro = int(input("Primeiro intervalo: "))
            segundo = int(input("Segundo intervalo: "))
            if segundo - primeiro <= 1000:
                for i in range(primeiro - 1, segundo):
                    print(f"{material[i]} {quantidade[i]} {preco[i]} {valorTotal[i]:.2%} {(individual[i]):.7%} {acumulada[i]:.7%} {classificacao[i]}")
            else:
                print(f"\nA diferença entre {primeiro} e {segundo} é de {segundo - primeiro} que é maior que 1000\n")
        case 5:
            plt.plot(acumulada)
            plt.show()
            #plt.bar(len(tabela), individual)
            #plt.show()
        case 6:
            abc = ['A', 'B', 'C']
            skus = propSKU()
            valor = propValor()
            listaCortes = [f'{corte[0]:.2%}', f'{corte[1]:.2%}', f'{corte[2]:.2%}']
            listaSKUs = [f'{skus[0]:.2%}', f'{skus[1]:.2%}', f'{skus[2]:.2%}']
            listaValores = [f'{valor[0]:.2%}', f'{valor[1]:.2%}', f'{valor[2]:.2%}']
            for i in range(len(individual)):
                individual[i] = f"{individual[i]:.2%}"
            for i in range(len(acumulada)):
                acumulada[i] = f"{acumulada[i]:.2%}"
            while len(listaCortes) != len(material):
                listaCortes.append('')
                listaSKUs.append('')
                listaValores.append('')
                abc.append('')
            data = {'Material': material,
                    'Quantidade': quantidade,
                    'Preço': preco,
                    'Valor Total': valorTotal,
                    'Individual': individual,
                    'Acumulada': acumulada,
                    'Classificação': classificacao,
                    '': '',
                    'Classificaçao': abc,
                    'Corte: ': listaCortes,
                    'Proporção de SKUs': listaSKUs,
                    'Proporção de Valor': listaValores
                    }
            nome = input("Digite o nome do Arquivo: ")
            tabela = pandas.DataFrame(data)
            tabela.to_excel(f'{nome}.xlsx', index = False)
            abrir = input("Quer abrir o arquivo[s/n]? ")
            if abrir == 's' or abrir == 'S':
                comando = ['cmd', '/c', 'start', '', f'./{nome}.xlsx']
                subprocess.run(comando, check=True)
        case 7:
            break
        case _:
            print("Opção Invalida")
