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

#DEF_Exercicio109: Modifique as funções que form criadas no desafio 107 para que elas aceitem um parâmetro a mais, 
#informando se o valor retornado por elas vai ser ou não formatado pela função moeda(), desenvolvida no desafio 108.

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
    return f'\033[0;34m{moeda}{num:.2f}\033[m'.replace('.',',')

#Aqui alteramos de forma booleana os parametros das defs, para apresentar ou não o valor formatado
#Deixando como padrão os valores não formatados