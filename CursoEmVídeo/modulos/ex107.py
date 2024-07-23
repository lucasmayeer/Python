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

#DEF_Exercicio107: Crie um módulo chamado moeda.py que tenha as funções incorporadas aumentar(), diminuir(), dobro() e metade(). 
#Faça também um programa que importe esse módulo e use algumas dessas funções.

def metade(num):
    metade = num / 2
    return f'\033[0;31m{metade}\033[m'


def dobro(num):
    dobro = num * 2
    return f'\033[0;31m{dobro}\033[m'


def diminuir(num, taxa):
    diminui = num - (taxa * num) / 100
    return f'\033[0;31m{diminui}\033[m'


def aumentar(num, taxa):
    aumento = num + (taxa * num) / 100
    return f'\033[0;31m{aumento}\033[m'