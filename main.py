import pandas
import subprocess
from tabulate import tabulate
from functions import *

#Coletar dados da Tabela
while True:
  caminho = input("Digite o nome do arquivo: ")
  caminho += ".xlsx"
  if os.path.exists(caminho):
    tabela = pandas.read_excel(caminho)
    material = tabela['Material'].tolist()
    quantidade = tabela['Quantidade'].tolist()
    preco = tabela['Preço'].tolist()
    valorTotal = tabela['Valor Total'].tolist()
    corte = [0.8, 0.95, 1]
    break
  else:
    print("Arquivo não existe")

#Eliminar Repetidos
#eliminarRepetidos(tabela, material, quantidade, preco, valorTotal)

#Cacular
total = calcularTotal(valorTotal)
individual = calcularIndividual(valorTotal, total)
acumulada = calcularaAcumulada(individual)
classificacao = calcularClassificacao(corte, acumulada)

#Menu
while True:
  op = int(input("\n1 - Modificar Corte\n2 - Proporção de SKUs\n3 - Proporção de Valor\n4 - Imprimir Tabela\n5 - Gerar Grafico\n6 - Exportar Tabela\n7 - Sair\nEscolha uma das opções acima: "))
  apagarTela()
  match op:
    case 1:
      corte[0] = float(input("Corte A: ")) / 100
      corte[1] = float(input("Corte B: ")) / 100
      corte[2] = float(input("Corte C: ")) / 100
      classificacao = calcularClassificacao(corte, acumulada)
    case 2:
      a, b, c = propSKU(classificacao)
      print(f"A: {a:.2%}\nB: {b:.2%}\nC: {c:.2%}")
    case 3:
      a, b, c = propValor(classificacao, individual)
      print(f"A: {a:.2%}\nB: {b:.2%}\nC: {c:.2%}")
    case 4:
      dados = []
      if len(tabela) > 500:
        print("Digite um intevalo para mostra os dados, sendo no máximo 500")
        primeiro = int(input("Primeiro intervalo: "))
        segundo = int(input("Segundo intervalo: "))
        if segundo - primeiro <= 500:
          for i in range(primeiro - 1, segundo):
            dados.append([material[i], quantidade[i], preco[i], valorTotal[i]])
        else:
          print(f"\nA diferença entre {primeiro} e {segundo} é de {segundo - primeiro} que é maior que 1000\n")
      else:
        for i in range(len(tabela)):
          dados.append([material[i], quantidade[i], preco[i], valorTotal[i]])
      tabular = tabulate(dados, ['Material', 'Quantidade', 'Preço', 'Valor Total'], tablefmt="grid", stralign="left")
      print(tabular)
    case 5:
      gerarGraficos(individual, acumulada)
    case 6:
      dados = gerarTabela(individual, acumulada, classificacao, corte, material, quantidade, preco, valorTotal)
      tabela = pandas.DataFrame(dados)
      tabela.style.set_properties(**{'font-family': 'Arial', 'color': 'blue'})
      nome = input("Digite o nome do Arquivo: ")
      tabela.to_excel(f'{nome}.xlsx', index = False)
      abrir = input("Quer abrir o arquivo[s/n]? ")
      if abrir == 's' or abrir == 'S':
        comando = ['cmd', '/c', 'start', '', f'./{nome}.xlsx']
        subprocess.run(comando, check=True)
    case 7:
      break
    case _:
      print("Opção Invalida")
