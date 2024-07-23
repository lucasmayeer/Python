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

""" TRATAMENTO DE ERROS
    Aqui teremos o tratamento de erros nas DEF, que nada mais é que um retorno personalizado das mensangens de erro do programa quando ocorrem;
    Neste esquema, o programa não necessariamente interrompe por definitivo, quando ocorre o erro. Aqui podemos tratar o erro e personalizar conforma 
    a demanda.
    sintaxe:
    try: 
        --- Basicamente é o teste da situação, tente executar isso(programa)...
    except(): 
        --- Aqui começa os tratamentos necessariamente, podemos ter varios excepts (dependendo da situação). Aqui testamos os erros, isto é, se acontecer erro X, faça isso:
        Temos uma vasta gama de erros, no qual podemos conciliar eles em except; EXEMPLO:
            - (ValueError, TypeError): Erros de digitação, podemos agrupar eles em um EXCEPT e retornar "ERRO! VALOR INFORMADO INVALIDO"
        Uma boa utilização para o except é que podemos fazer ele retornar para o inicio (SE FOR UM LAÇO) fazendo assim com que o ERRO do programa,
        não interrompa seu funcionamento
    else:
    -- Aqui é a finalização, se tudo estiver certo e não tiver nenhum erro: faça isso! (aqui normalmente se coloca o resultado do seu try)
"""
#DEF_Exercicio113:Reescreva a função leiaInt() que fizemos no desafio 104, 
#incluindo agora a possibilidade da digitação de um número de tipo inválido. Aproveite e crie também uma função leiaFloat() com a mesma funcionalidade.

def leiaInt(msg):
    while True:
        try: #Tente isso
            n = int(input(msg))
        except (ValueError, TypeError): #Se der erro de de valor invalido ou tipo de valor invalido, faça isso:
            print(f'\033[0;31mERRO! Digite um número INTEIRO válido.\033[m') #Execute a mensagem de erro
            continue #E volte para o inicio do laço
        except (KeyboardInterrupt): #Se der erro de interrupção do programa pelo usuario, faça isso:
            print(f'\033[0;31mEntrada de dados interrompida pelo usuario.\033[m') #execute a mensagem de erro
            return 0 #finalize o programa retornando 0. (por que o usuario quis interromper o programa)
        else: #se nada der erro:
            return n #retorne com o resultado do input


def leiaFloat(msg):
    while True:
        try:
            n = float(input(msg))
        except (ValueError, TypeError):
            print(f'\033[0;31mERRO! Digite um número REAL válido.\033[m')
            continue
        except (KeyboardInterrupt):
            print(f'\033[0;31mEntrada de dados interrompida pelo usuario.\033[m')
            return 0
        else:
            return n        

