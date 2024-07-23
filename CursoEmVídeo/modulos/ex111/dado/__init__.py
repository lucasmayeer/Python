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

#DEF_Exercicio110: Adicione o módulo moeda.py criado nos desafios anteriores, 
#uma função chamada resumo(), que mostre na tela algumas informações geradas pelas funções que já temos no módulo criado até aqui.

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
    return f'\033[0;35m{moeda}{num:.2f}\033[m'.replace('.',',')


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

#Feito um resumo do codigo, de forma tabular - com a DEF RESUMO();
# o comando \t serve para aplicar uma tabulação na string; podemos aplicar varias tabulacoes ate encaixar como deseja EX: \t\t\t{moeda(num)}
