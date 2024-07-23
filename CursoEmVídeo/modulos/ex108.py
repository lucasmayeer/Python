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

#DEF_Exercicio108: Adapte o código do desafio #107, criando uma função adicional chamada moeda() que consiga mostrar os números como um valor monetário formatado.

def metade(num=0):
    metade = num / 2
    return metade


def dobro(num=0):
    dobro = num * 2
    return dobro


def diminuir(num=0, taxa=0):
    diminui = num - (taxa * num) / 100
    return diminui


def aumentar(num=0, taxa=0):
    aumento = num + (taxa * num) / 100
    return aumento


def moeda(num=0, moeda='R$'):
    return f'\033[0;32m{moeda}{num:.2f}\033[m'.replace('.',',')

#replace(' . ' , ' , ') - Estou pedindo para trocar todos os pontos (.) por virgulas (,)
#Deixamos todos os valores das DEF opcionais com o sinal de igual

#Aqui, so temos retorno da def MOEDA, as demais so servem como base de calculo;
#Por isso a unica def que retorna uma STRING formatada é a moeda()



