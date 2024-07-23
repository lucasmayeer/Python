"""AMBIENTE DE MODULARIZAÇÃO:
    A modularização em Python é o processo de dividir um programa em módulos ou partes menores, cada uma com uma função específica. 
    Isso é importante porque torna o código mais organizado, fácil de entender e manter. 
    Além disso, permite reutilizar código em diferentes partes do programa ou até mesmo em diferentes projetos. 
    Com a modularização, você pode isolar partes do seu código, o que facilita a identificação e correção de erros, 
    além de promover a colaboração em equipe de forma mais eficiente.

    VANTAGENS:
    - Organização do código;
    - Facilidadse na manutenção;
    - Ocultação de código detalhado;
    - Reutilização em outros projetos;
"""

#DEF_Exercicio112: Dentro do pacote utilidadesCeV que criamos no desafio 111, temos um módulo chamado dado. 
#Crie uma função chamada leiaDinheiro() que seja capaz de funcionar como a função imputa(), mas com uma validação de dados para aceitar apenas valores que seja monetários.

def metade(num=0, formato=False):
    metade = num / 2
    return metade if formato is False else moeda(metade)


def dobro(num=0, formato=False):
    dobro = num * 2
    return dobro if formato is False else moeda(dobro)


def diminuir(num=0, taxa=0, formato=False):
    diminui = num - (taxa * num) / 100
    return diminui if formato is False else moeda(diminui)


def aumentar(num=0, taxa=0, formato=False):
    aumento = num + (taxa * num) / 100
    return aumento if formato is False else moeda(aumento)


def moeda(num=0, moeda='R$'):
    return f'\033[0;38m{moeda}{num:.2f}\033[m'.replace('.',',')


def resumo(num=0, aumento=10, reducao=5):
    print('-'*34)
    print(f'{"RESUMO DO VALOR":^34}')
    print('-'*34)

    print(f'Preço analisado: \t{moeda(num)}')
    print(f'Dobro do preço: \t{dobro(num,True)}')
    print(f'Metade do preço: \t{metade(num,True)}')
    print(f'{aumento}% de aumento: \t{aumentar(num,aumento,True)}')
    print(f'{reducao}% de aumento: \t{diminuir(num,reducao,True)}')

    print('-'*34)

def leiaDinheiro(msg):
    valido = False

    while not valido:
        entrada = str(input(msg)).replace(',' , '.').strip() #Strip ignora os espaços no input - replace troca as virgulas (caso usadas) por pontos na hora do input
        if entrada.isalpha() or entrada == '': #Se o input for alfanumerico (ou seja, numero + letras) OU se o valor do input estiver em branco
            print(f'\033[0;31mERRO: \"{entrada}\" é um preço inválido!\033[m') #MSG DE ERRO
        else:
            valido = True #Quebra do loop
            return float(entrada) #Converte o valor informado no input para float. (Depois de toda a validação e tratamento do IF)
        
#A def leiaDinheiro() serve para basicamente validar os dados postos;
#Nela, caso seja colocado valores em letras __isalpha()__ retorna mensagem de erro;
#Nela, caso o usuario coloque , (virgula) como separador de real p/ centavo, convertemos para o python para .(ponto) __.replace(',' , '.')__
        #OBS: COMO no input de entrada É CONSIDERADO UMA STRING, podemos utilizar o replace para substituir. Essa DEF não soma nenhum valor;
        # então não precisamos converter os numeros para int ou float, por isso deixamos em STR por que podemos malear os dados com melhor precisão.
#Nela, ignoramos os espaços que o usuario coloca no input __.strip()__
#Nela, se o usuario deixar o input entrada em branco, também da mensagem de erro __or entrada == '':__       
           
#Se o valor for numerico (isnumeric), os espaços foram excluidos (strip()), e indenpende se foi separado por virgula ou ponto, se tudo foi concluido;
        #O else de retorno retorna o valor inserido no input, AI SIM CONVERTENDO O VALOR INFORMADO PARA FLOAT()