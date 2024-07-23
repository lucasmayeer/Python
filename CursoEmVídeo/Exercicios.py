#Exercicio001
ex001 = 'Olá, Mundo'
print(ex001)

#Exercicio002
nome = input('Qual o seu nome?')
print('É um prazer te conhecer',nome)

#Exercicio003
n1 = int(input('Digite um numero'))
n2 = int(input('Digite um numero'))
soma = n1 + n2

print('A soma entre {} e {} resulta em {}'.format(n1,n2,soma))

#Exercicio004
n = input('Digite algo:')

print('O tipo primitivo desse valor é:',type(n))
print('So tem espaços?',n.isspace())
print('É um numero?',n.isnumeric())
print('É alfabetico?',n.isalpha())
print('É alfanumerico?',n.isalnum())
print('Está em maisculas?',n.isupper())
print('Está em minusculas?',n.islower())
print('É alfanumerico?',n.istitle())

#Exercicio005
n = int(input('Digite um numero: '))

print('{} Seu sucessor é: {} e seu antecessor é {}'.format(n,n+1,n-1))

#Exercicio006
n = int(input('Digite um numero: '))

print('Valor dobrado: {}'.format(n*2))
print('Valor triplicado: {}'.format(n*3))
print('Raiz quadrada: {:.2f}'.format(n**(1/2)))

#Exercicio007
n1 = float(input('Insira a primeira nota: '))
n2 = float(input('Insira a segunda nota: '))
media = (n1+n2)/2

print('A média entre {:.1f} e {:.1f} é {:.1f}'.format(n1,n2,media))

#Exercicio008
m = float(input('Insira os metros: '))

print('O valor {}, convertido em centimetros fica: {:.0f}cm'.format(m,(m*100)))
print('O valor {}, convertido em milimetros fica: {:.0f}mm'.format(m,(m*1000)))

#Exercicio009
n = int(input('Insira um numero: '))

print('-'*10)
print('{} x  1 = {}'.format(n,n*1))
print('{} x  2 = {}'.format(n,n*2))
print('{} x  3 = {}'.format(n,n*3))
print('{} x  4 = {}'.format(n,n*4))
print('{} x  5 = {}'.format(n,n*5))
print('{} x  6 = {}'.format(n,n*6))
print('{} x  7 = {}'.format(n,n*7))
print('{} x  8 = {}'.format(n,n*8))
print('{} x  9 = {}'.format(n,n*9))
print('{} x 10 = {}'.format(n,n*10))
print('-'*10)

#Exercicio010
c = float(input('Informe sua renda: R$'))
d = float(3.27)

print('Com o valor de R${} você pode comprar US${:.2f}'.format(c,(c/d)))

#Exercicio011
a = float(input('Informa a altura da parede: '))
l = float(input('Informa a largura da parece: '))
d = l*a
t = 2

print('A área de sua parede é {}m2, será necessario {}l de tinta para pintura'.format(d,d/t))

#Exercicio012
p = float(input('Insira o valor do produto: R$'))
d = (5*p)/100
print('O valor R${:.2f} com 5% de desconto fica: R${:.2f}'.format(p,p-d))


#Exercicio013
s = float(input('Insira seu salario: '))
a= (15*s)/100

print('O salario R${} com um 15% de aumento fica: R${}'.format(s,s+a))


#Exercicio014
c = float(input('Qual a temperatura em Celsius: '))
f = (c*9/5)+32

print('A temperatura de {}°C corresponde á {}°F'.format(c,f))

#Exercicio015
a = int(input('Quantos dias alugado:'))
km = float(input('Quantos KM rodados:'))
total = (a*60) + (km*0.15)

print('O total a pagar é de R${:.2f}'.format(total))


#Exercicio016
#PRIMEIRO JEITO DE RESOLVER (COM BIBLIOTECAS)

from math import trunc

n1 = float(input('Informe um número real:'))

print('O número {} tem a parte inteira {}'.format(n1,trunc(n1)))
#------------------------------------------------------------------------

#SEGUNDO JEITO DE RESOLVER (SEM BIBLIOTECAS)
n1 = float(input('Informe um número real:'))

print('O número {} tem a parte inteira {}'.format(n1,int(n1)))


#Exercicio017
#PRIMEIRO JEITO DE RESOLVER (COM BIBLIOTECAS)

from math import hypot

n1= float(input('Qual o comprimento do Cateto Oposto: '))
n2= float(input('Qual o comprimento do Cateto Adjacente: '))

print('A Hipotenusa vai medir {:.2f}'.format(hypot(n1,n2)))
#------------------------------------------------------------------------
#SEGUNDO JEITO DE RESOLVER (SEM BIBLIOTECAS)
co = float(input('Qual o comprimento do Cateto Oposto: '))
ca = float(input('Qual o comprimento do Cateto Adjacente: '))
hip = (co **2 + ca ** 2) ** (1/2)

print('A Hipotenusa vai medir {:.2f}'.format(hip))


#Exercicio018
import math

n1 = float(input('Informa um Ângulo: '))
sen = math.sin(math.radians(n1))
cos = math.cos(math.radians(n1))
tan = math.tan(math.radians(n1))

print('O Ângulo de {}, tem seno de {:.2f}, cosseno de {:.2f} e a tangente de {:.2f}'.format(n1,sen,cos,tan))

#Exercicio019
import random

a1 = str(input('Informe o primeiro aluno: '))
a2 = str(input('Informe o segundo aluno: '))
a3 = str(input('Informe o terceiro aluno: '))
a4 = str(input('Informe o quarto aluno: '))
lista = [a1,a2,a3,a4]

print(random.choice(lista))


#Exercicio020
import random

a1 = str(input('Informe o primeiro aluno: '))
a2 = str(input('Informe o segundo aluno: '))
a3 = str(input('Informe o terceiro aluno: '))
a4 = str(input('Informe o quarto aluno: '))
lista = [a1,a2,a3,a4]
random.shuffle(lista)

print('A ordem de apresentação será:' )
print(lista)


#Exercicio021
#import pygame (importando a biblioteca de jogos para rodar musica)
#pygame.init() (Iniciando a biblioteca)
#pygame.mixe.music.load('arquivo.mp3') (Escolhendo o Arquivo MP3)
#pygame.mixer.music.play() (Iniciando o play)
#pygame.event.wait() (Para encerrar o programa, ele vai ter que esperar a musica rodar)


#Exercicio022
nome = str(input('Digite seu nome COMPLETO: ')).strip()
pnome = nome.split()

print('Convertido em todas as letras maísculas: {}'.format(nome.upper()))
print('Convertido em todas as letras maísculas: {}'.format(nome.lower()))
print('Seu nome ao todo possui {} letras (sem considerar espaços)'.format(len(nome) - nome.count(' '))) #EM (len(nome) - nome.count(' ')) contamos o tamanho da STRING nome e substraimos (-) com o contador de ESPAÇOS (' ')
print('Seu primeiro nome é {}, e possui {} letras'.format(pnome[0],len(pnome[0])))

#Exercicio023
num = int(input('Insira um numero:'))

print('Analisando o número {} informado...'.format(num))
print('Unidade: {}'.format(num // 1 % 10)) 
print('Dezena: {}'.format(num // 10 % 10)) 
print('Centena: {}'.format(num // 100 % 10)) 
print('Milhar: {}'.format(num // 1000 % 10))

#Como no FORMAT:
    # Temos Unidade: o Retorno da conta ( num// 1 ) vai retornar duas casas, pegamos o resto da divisão por % 10 e extraimos a unidade(UNIDADE) que restou 
    # Temos Dezena: o Retorno da conta ( num // 10 ) vai retornar duas casas, pegamos o resto da divisão por % 10 e extraimos a unidade (DEZENA) que restou 
    # Temos Centena: o Retorno da conta ( num // 100 ) vai retornar duas casas, pegamos o resto da divisão por % 10 e extraimos a unidade (CENTENA) que restou 
    # Temos Milhar: o Retorno da conta ( num // 1000 ) vai retornar duas casas, pegamos o resto da divisão por % 10 e extraimos a unidade (MILHAR) que restou 


#Exercicio024
cidade = str(input('Qual cidade você nasceu?')).strip()

print(cidade[:5].upper() == 'SANTO')

# Explicando a SINTAXE:
    # .strip() elimina os possiveis espaços dados no INPUT (Espaços antes e depois da palavra) --> EX: '   Curitiba   '
    # Sobre o PRINT: 
        # (cidade[:5]) --> Aqui indica que quero trazer os caracteres [0 até (:) 05]. Quando não informado o numero antes dos [:], é indexado em 0.
        # (.upper) --> O (.upper) vai SEMPRE converter o resultado de cidade em letras maisculas. EX: Se escrever 'Curitiba', o retorno da analise converterá para 'CURITIBA'
        # (== 'SANTO') --> Aqui aplicamos o resultado do desafio, que pedia o retorno se o INPUT(Usuario) nasceu em alguma cidade que contenha "SANTO"...
            # a conversão em UPPER permite que o usuario na hora do INPUT escreva de qualquer forma, pois será indiferente o resultado da analise (== 'SANTO')
            # uma vez que convertemos tudo em LETRAS MAISCULAS com o .upper.


#Exercicio025
nome = str(input('Qual o seu nome?')).strip()

print('Seu nome tem Silva ? {}'.format('silva' in nome.lower()))

# Explicando a SINTAXE (O DESAFIO ERA RETORNAR {TRUE} OR {FALSE} EM CASO DO NOME INFORMADO CONTER 'SILVA')
    # (.strip()) --> Elimina os possiveis espaços dados no INPUT (Espaços antes e depois da palavra) --> EX: '   Curitiba   '
    # ('silva' in nome.lower())--> Verifica se TEM 'silva' no INPUT informado(nome). E convertemos qualquer INPUT em letras minisculas, para que a analise seja feita em minusculas
        # e deixe indiferente a forma como escreveram no INPUT


#Exercicio026
frase = str(input('Digite uma Frase:')).upper().strip()

print('A letra A aparece {} na frase.'.format((frase.count('A'))))
print('A primeira letra A apareceu na posição {}.'.format(frase.find('A') + 1))
print('A última letra A apareceu na posição {}.'.format(frase.rfind('A') + 1 ))

# Explicando a SINTAXE (O DESAFIO ERA RETORNAR 3 PRINTS (1. Quantas vezes a letra 'A' se repete), (2. Qual a posição do primeiro 'A'), (3. Qual a posição do ultimo 'A'))
    # (.upper().strip()) --> Já no INPUT(frase), determinamos que o retorno em sistema será TODAS AS LETRAS MAISCULAS (.upper()) e que será ignorado na CONTAGEM os espaços (.strip())
    # (frase.count('A')) --> Aqui, como deixamos o retorno do INPUT em maisculo (.upper()), iremos calcular quantas vezes se repete o 'A' em maisculo;
    # (frase.find('A') + 1) --> Aqui utilizamos o .find, para encontrar o PRIMEIRO 'A' da frase. No caso, EM SISTEMA é indexado em 0, para o retorno do usuario NÃO. 
        # Por isso localizamos a primeira posição de 'A' e somamos ( + 1) para dar o retorno correto no INPUT. EX: Arara (Para gente, o A está em primeiro, em sistema ele considera na posição 0)
    # (frase.rfind('A') + 1 ) --> Aqui é a mesma coisa, so altera que localizamos o ULTIMO 'A' da frase. Neste caso utilizamos .rfind('A'), que vai analisar o INPUT da DIREITA para a ESQUERDA.
        # então para o PYTHON ele vai ir atras do primeiro 'A' so que da direita para a esquerda, porém ele vai respeitar as posições padrões do sistema (ESQUERDA 0,1,2,3,4,5...)


#Exercicio027
nome1 = str(input('Digite seu nome:')).strip()
nome2 = nome1.split()

print('Seu primeiro nome é {}'.format(nome2[0]))
print('Seu último nome é {}'.format(nome2[len(nome2)-1]))

# Explicando a SINTAXE (O DESAFIO ERA ERA RETORNAR O PRIMEIRO NOME E ULTIMO NOME do Input)
    # .strip() --> Elimina os possiveis espaços dados no INPUT (Espaços antes e depois da palavra) --> EX: '   Curitiba   '
    # .split() --> Quebramos em uma lista o nome informado no input, aqui toda vez que encontra um valor espaçado, ele ''RESETA'' as posições de caracteres
        # EX: [Lucas, Mayer] Aqui com o .split() L está na posição 0 e M também está na posição 0, pois são considerados diferentes.E também temos uma ordem das palavras como um todo
        # então [Lucas] está na posição 0 e [Mayer] está na posição 1. Então temos posições entre palavras, e dentro das palavras temos as posições dos caracteres.
    # .format(nome2[0]) --> Trazemos a PALAVRA que está na posição ZERO. Ou seja, para o usuario, seria o PRIMEIRO NOME informado;
    # .format(nome2[len(nome2)-1]) --> Aqui também, porém a ideia é que utilizemos len(nome2) para sabermos quantas posições(PALAVRAS) TEM no nome2.
        # Com isso, como é indexado em [0] subtraimos -1 ([len(nome2)-1]) para que ele traga a ULTIMA POSIÇÃO de nome.  

#Exercicio028
import random

lista = [0,1,2,3,4,5]
random= random.choice(lista)
choice = int(input('Escolha um número de 0 até 5:'))

if choice == random:
    print('O número selecionado foi {}'.format(choice))
    print('O número selecionado pelo sistema foi {}'.format(random))
    print('Parabens, você venceu!!')
else:
    print('O número selecionado foi {}'.format(choice))
    print('O número selecionado pelo sistema foi {}'.format(random))
    print('Sinto muito, você perdeu!!')


#Exercicio029
velocidade = float(input('Informa a velocidade do veículo:'))
vpermitido = 80

if velocidade > vpermitido:
    print('Você recebeu uma multa por ultrapassar a velocidade de {}Km/h permitida para a via!'.format(vpermitido))
    print('Valor a pagar de R${:.2f}'.format((velocidade - vpermitido) * 7))
else:
    print(' ')

#Exercicio030
n = int(input('Me diga um número qualquer:'))
resultado = n % 2

if resultado == 0:
    print('O número {} é PAR'.format(n))
else:
    print('O número {} é ÍMPAR!'.format(n))    

# (resultado = n % 2) --> Aqui é um retorno matematico, sempre que esse resultado dar 0, matematicamente o numéro é PAR. se der 1 é IMPAR


#Exercicio031
viagem = float(input('Informa quantos KM terá sua viagem:'))

if viagem <= 200.00:
    print('Sua viagem tera {}Km'.format(viagem))
    print('O valor da viagem fica R${:.2F}'.format(viagem * 0.50))
    print('Cobrado o valor de R$0.50 por Km')
else:
    print('Sua viagem tera {}Km'.format(viagem))
    print('O valor da viagem fica R${:.2F}'.format(viagem * 0.45))
    print('Cobrado o valor de R$0.50 por Km')


#Exercicio032
from datetime import date

ano = int(input('Que ano quer analisar? Coloque 0 para analisar o ano atual: '))

if ano == 0:
    ano = date.today().year ##Aqui retira o ano do sistema, com datetime
if ano % 4 == 0 and ano % 100 != 0 or ano % 400 == 0: #Matematicamente o calculo para ver se um ano é BISEXTO ou não
    print('O Ano {} é BISEXTO'.format(ano))
else:
    print('O Ano {} não é BISEXTO'.format(ano))    


#Exercicio033
a = int(input('Primeiro valor:'))    
b = int(input('Segundo valor:'))    
c = int(input('Terceiro valor:'))    

menor = a #Declarando uma variavel ja como MENOR
maior = a #Declarando uma variavel ja como MAIOR

if b < a and b < c:
    menor = b
if c < a and c < b:
    menor = c
#Fazendo o teste logico do MENOR, se a variavel [menor = a] NÃO FOR de fato MENOR, ele altera a VARIAVEL dentro do IF para a variavel DE FATO menor

if b > a and b > c:
    maior = b
if c > a and c > b:
    maior = c
#Fazendo o teste logico do MAIOR, se a variavel [maior = a] NÃO FOR de fato MAIOR, ele altera a VARIAVEL dentro do IF para a variavel DE FATO maior

print('O menor valor digitado foi {}'.format(menor))
print('O maior valor digitado foi {}'.format(maior))


#Exercicio034
salario = float(input('Informe qual o seu salario: R$'))

if salario > 1250.00:
    print('O salario informado de R${:.2F} terá um aumento de 10%!'.format(salario))
    print('Valor do salario com aumento: R${:.2F}'.format((10*salario)/100 + salario))
else:
    print('O salario informado de R${:.2F} terá um aumento de 15%!'.format(salario))
    print('Valor do salario com aumento: R${:.2F}'.format((15*salario)/100 + salario))


#Exercicio035
r1 = float(input('Informe o primeiro segmento:'))
r2 = float(input('Informe o segundo segmento:'))
r3 = float(input('Informe o terceiro segmento:'))

if r1 < r2 + r3 and r2 < r1 + r3 and r3 < r1 + r2: ##Principio matematico para calcular se 3 linhas podem ou não formar um triangulo; 
    print('Os segmentos acima PODEM FORMAR um triângulo') ##Todos os segmentos precisam ser menor que a soma dos outros dois segmentos 
    print('Os segmentos acima PODEM FORMAR um triângulo')
else:
    print('Os segmentos acima NÃO FORMAM um triângulo')


#Exercicio036
casa= float(input('Qual o valor da Casa? R$ '))
salario= float(input('Qual o seu salario? R$'))
tempo= int(input('Você deseja parcelar em quantos anos?'))

prestacao= casa / (tempo*12)
salario30= (30*salario)/100

if prestacao <= salario30:
    print('A prestação mensal da sua casa ficará R${:.2f}'.format(prestacao))
    print('Dados do emprestimo')
    print('Valor da Casa: R${:.2f}'.format(casa))
    print('Salario Informado: R${:.2f}'.format(salario))
    print('Parcelado em {} Anos'.format(tempo))
elif prestacao > salario30:
    print('A prestação mensal da sua casa ficaria R${:.2f}'.format(prestacao))
    print('O valor da prestação ultrapassa 30% de seu salario {:.2f}'.format(salario30))
    print('EMPRESTIMO NEGADO')


#Exercicio037
num= int(input('Digite um numero inteiro:'))

print('''Escolha uma das bases para conversão:
      [ 1 ] Converter para BINARIO
      [ 2 ] Converter para OCTAL
      [ 3 ] Converter para HEXADECIMAL ''')

opcao = int(input('Sua opção:'))

if opcao == 1:
    print('{} Convertido para BINARIO é {}'.format(num,bin(num)[2:]))
elif opcao == 2:
    print('{} Convertido para OCTAL é {}'.format(num,oct(num)[2:]))
elif opcao == 3:
    print('{} Convertido para HEXADECIMAL é {}'.format(num,hex(num)[2:]))
else:
    print('A opção seleciona é INVALIDA!')


#Exercicio038
n1= int(input('Informe o primeiro valor:'))
n2= int(input('Informe o segundo valor:'))

if n1 > n2:
    print('O primeiro valor {} é MAIOR que o segundo valor {}'.format(n1,n2))
elif n2 > n1:
    print('O segundo valor {} é MAIOR que o primeiro valor {}'.format(n2,n1))
elif n1 == n2:
    print('Não existe valor maior, os dois números são iguais')

    
#Exercicio039
from datetime import date

anoatual = date.today().year
nasc = int(input('Ano de Nascimento:')) 
idade = anoatual - nasc  

print('Quem nasceu em {} tem {} anos em {}'.format(nasc,idade,anoatual))
if idade == 18:
    print('Você tem que se alistar IMEDIATAMENTE!')
elif idade < 18:
    saldo = 18 - idade
    ano = anoatual + saldo
    print('Ainda faltam {} anos para o alistamento'.format(saldo))
    print('Seu alistamento será em {}'.format(ano))
elif idade > 18:
    saldo = idade - 18
    ano = anoatual - saldo
    print('Você já deveria ter se alistado há {} anos.'.format(saldo))
    print('Seu alistamento foi em {}'.format(ano))


#Exercicio040
nota1 = float(input('Digite sua primeira nota:')) 
nota2 = float(input('Digite sua segunda nota:'))
media = (nota1 + nota2) / 2

if media < 5:
    print('REPROVADO!')
    print('Sua média foi {:.1f}'.format(media))
elif media >= 5 and media < 7:
    print('RECUPERAÇÃO!')
    print('Sua média foi {:.1f}'.format(media))
elif media >= 7 :
    print('APROVADO!')
    print('Sua média foi {:.1f}'.format(media))


#Exercicio041
idade = int(input('Qual a sua idade?'))

if idade < 10:
    print('Pela sua idade de {}, você se classifica como nadador MIRIM.'.format(idade))
elif idade < 15:
    print('Pela sua idade de {}, você se classifica como nadador INFANTIL.'.format(idade))
elif idade < 20 :
    print('Pela sua idade de {}, você se classifica como nadador JUNIOR.'.format(idade))
elif idade < 26:
    print('Pela sua idade de {}, você se classifica como nadador SÊNIOR.'.format(idade))
else:
    print('Pela sua idade de {}, você se classifica como nadador MASTER.'.format(idade))

    
#Exercicio042
r1 = float(input('Informe o primeiro segmento:'))
r2 = float(input('Informe o segundo segmento:'))
r3 = float(input('Informe o terceiro segmento:'))

if r1 < r2 + r3 and r2 < r1 + r3 and r3 < r1 + r2: ##Principio matematico para calcular se 3 linhas podem ou não formar um triangulo; 
    print('Os segmentos acima PODEM FORMAR um ', end='') ##Todos os segmentos precisam ser menor que a soma dos outros dois segmentos
    if r1 == r2 == r3: ##Principio matematico
        print('TRIANGULO EQUILATERO!')
    elif r1 != r2 != r3 != r1: ##Principio matematico
        print('TRIANGULO ESCALENO!')
    else:
        print('TRIANGULO ISOSCELES!')
else:
    print('Os segmentos acima NÃO FORMAM um triângulo')


#Exercicio043
peso = float(input('Qual o seu peso? (Kg)'))
altura = float(input('Qual a sua altura? (m)'))
imc = peso / (altura ** 2)

if imc < 18.5:
    print('Abaixo do Peso! - IMC DE: {:.2f}Kg/m2'.format(imc))
elif imc >= 18.5 and imc <= 25:
    print('Peso Ideal! - IMC DE: {:.2f}Kg/m2'.format(imc))
elif imc >= 25 and imc <= 30:   
    print('Sobrepeso! - IMC DE: {:.2f}Kg/m2'.format(imc))
elif imc >= 30 and imc <= 40:
    print('Obesidade! - IMC DE: {:.2f}Kg/m2'.format(imc))
elif imc > 40:
    print('Obesidade Mórbida! - IMC DE: {:.2f}Kg/m2'.format(imc))   
    
    
#Exercicio044
preco = float(input('Informe o valor das compras: R$ '))
print(''' FORMAS DE PAGAMENTOS
      [ 1 ] Á vista no PIX/Dinheiro
      [ 2 ] Á vista no cartão
      [ 3 ] 2x no cartão SEM JUROS
      [ 4 ] 3x ou mais no cartão COM JUROS''')

opcao = int(input('Qual é a opção? '))

if opcao == 1:
    total = preco - (preco * 10 ) / 100
elif opcao == 2:
    total = preco - ( preco * 5 ) / 100
elif opcao == 3:
    total = preco
    parcela = total / 2
    print('Sua compra será parcelada em 2x de R${:.2F} SEM JUROS'.format(parcela))
elif opcao ==4:
    total = preco + (preco * 20) / 100
    totalparcelas = int(input('Quantas parcelas? '))
    parcela = total / totalparcelas
    print('Sua compra será parcelada em {}x de R${:.2F} COM JUROS'.format(totalparcelas,parcela))
else:
    print('FORMA DE PAGAMENTO INVALIDA!')

print('Sua compra de R${:.2F} vai custar R${:.2F} no final.'.format(preco,total))


#Exercicio045
from random import randint

#Variaveis JOKENPO - COMPUTADOR
itens = ('Pedra', 'Papel', 'Tesoura')
computador = randint(0,2)

#Variaveis JOKENPO - JOGADOR
print(''' Escolha:
      [ 0 ] Pedra
      [ 1 ] Papel
      [ 2 ] Tesoura''')
jogador = int(input('Qual é a sua jogada?'))

#Resultado da jogada
print('-=' * 11)
print('O computador jogou {}'.format(itens[computador]))
print('Jogador jogou {}'.format(itens[jogador]))
print('-=' * 11)


#DETERMINANDO QUEM VENCEU
if computador == 0: # computador jogou pedra
    if jogador == 0:
        print('Empate')
    elif jogador == 1:
        print('Jogador Venceu!!')
    elif jogador == 2:
        print('Computador Venceu!!')
    else:
        print('Jogada Invalida! Tente Novamente')
elif computador == 1: # computador jogou papel
    if jogador == 0:
        print('Computador Venceu!!')
    elif jogador == 1:
        print('Empate')
    elif jogador == 2:
        print('Jogador Venceu!!')
    else:
        print('Jogada Invalida! Tente Novamente')
elif computador == 2: # computador jogou tesoura
    if jogador == 0:
        print('Jogador Venceu!!')
    elif jogador == 1:
        print('Computador Venceu!!')
    elif jogador == 2:
        print('Empate')
    else:
        print('Jogada Invalida! Tente Novamente')


#Exercicio046
from time import sleep

print('Contagem regressiva para a queima de fogos:')        
for c in range (10, -1, -1):
    print(c)
    sleep(1)
print('FELIZ ANO NOVO!') 


#Exercicio047
print('Esse programa vai te mostrar todos os números PARES de 01 até 50:')
for c in range (2, 51, 2):
    print(c, end = ' ')


#Exercicio048
soma = 0
cont = 0

for c in range (1,501,2):
    if c % 3 == 0:
        cont += 1 
        soma += c 
        
print('A soma de todos os {} valores solicitados é {}'.format(cont,soma))


#Exercicio049
print('TABUADA')
n = int(input('Insira um numero: '))

print('-'*10)
for c in range(1,11):
    print('{} x {:2} = {}'.format(n,c,n*c))
print('-'*10)


#Exercicio050

soma = 0
cont = 0
for c in range (1,7):
    num = int(input('Digite o {} valor: '.format(c)))
    if num % 2 == 0:
        soma += num
        cont += 1 

print('Você informou {} números PARES e a soma foi de {}'.format(cont,soma))    


#Exercicio051
print('Vamos calcular Progressão Aritmetica')
n = int(input('Insira um numero: '))
pa = int(input('Insira a progressão: '))
decimo = n + (10-1) * pa

for c in range (n,decimo + pa ,pa):
    print('{}'.format(c), end=' -> ')

print('FIM!')


#Exercicio052
n = int(input('Insira um número: '))
totaldivisivel = 0

for c in range (1,n+1):
    if n % c == 0: # Calculando se os numéros são divisiveis por ele mesmo!
        print('\033[33m', end=' ') #Destacando quais eles são divisiveis
        totaldivisivel += 1 # Adicionando a contagem a QTD de divisões
    else:
        print('\033[31m', end=' ')

    print('{}'.format(c), end=' ')

print('\n\033[mO número {} foi divisivel {} vezes'.format(n,totaldivisivel)) 

if totaldivisivel == 2: 
    print('E por isso ele É PRIMO')
else:
    print('E por isso ele NÃO É PRIMO')


#Exercicio053
frase = str(input('Digite uma frase: ')).strip().upper()
palavras = frase.split()
palavrasjunto = ''.join(palavras)
inverso = ''

## inverso = [::-1] ----> Outra forma se fazer o exercicio, aqui essa variavel substitui o FOR! {{{MACETE DO FATIAMENTO}}}

for letra in range (len(palavrasjunto) - 1, -1, -1): #Aqui ele inverteu a palavra digitada na FRASE, e foi de traz para a frente. {{FRASE: Lucas -- INVERTO: sacuL}}
    inverso += palavrasjunto[letra]

print('O inverso de {} é {}'.format(palavrasjunto,inverso))

if inverso == palavrasjunto:
    print('Temos um palindromo')
else:
    print('A frase digitada não é um palindromo')


#Exercicio054
from datetime import date
anoatual = date.today().year

simmaioridade = 0
naomaioridade = 0

for ano in range (1,8):
    nasc = int(input('Em que ano a {}ª pessoa nasceu?'.format(ano)))  
    idade = anoatual - nasc

    if idade >= 21:
        simmaioridade += 1 
    else:
        naomaioridade += 1 

print('Ao todo tivemos {} pessoas maiores de idade'.format(simmaioridade))        
print('E também tivemos {} pessoas menores de idade!'.format(naomaioridade))          


#Exercicio055

maiorp = 0
menorp = 0

for p in range (1,6):
    peso = float(input('Peso da {}ª pessoa:'.format(p)))
    if p == 1: 
        maiorp = peso
        menorp = peso
#Esse IF, é a primeira ocorrencia dentro da variavel -- Isso é, a primeira ocorrencia (ou primeiro insert) 
#sempre vai se caracterizar como o [maiorp] e o [menorp], por que não tem as demais ocorrencias para comparar
    else:
        if peso > maiorp:
            maiorp = peso
        if peso < menorp:
             menorp = peso
#Nesses dois ultimos IF, ele analisa cada ocorrencia com a primeira, e se der TRUE, ele altera as variaveis [menorp] ou [maiorp]

print('O maior peso lido foi de {}Kg'.format(maiorp))        
print('O menor peso lido foi de {}Kg'.format(menorp))       

#Exercicio056

#Variaveis IDADE
somaidade = 0
mediaidade = 0

#Variaveis Nome e idade do HOMEM MAIS VELHO
idadeH = 0
nomeH = ''

#Variaveis Mulheres ACIMA de 20 anos
mulher20 = 0

for p in range (1,5):
    print('----- {}ª PESSOA -----'.format(p))
    nome = str(input('Nome: ')).strip()
    idade = int(input('Idade: '))
    sexo = str(input('Sexo [M/F]: ')).strip()

    somaidade += idade #Alterando a variavel da idade para os dados do input

    if p == 1 and sexo in 'Mm':
        idadeH = idade
        nomeH = nome
        #Esse IF, é a primeira ocorrencia dentro da variavel -- Isso é, a primeira ocorrencia (ou primeiro insert) 
        #sempre vai se caracterizar como o [idadeH] e o [nomeH], por que não tem as demais ocorrencias para comparar
    if sexo in 'Mm' and idade > idadeH:
        idadeH = idade
        nomeH = nome  
        #Nesses dois ultimos IF, ele analisa cada ocorrencia com a primeira, e se der TRUE, ele altera as variaveis [idadeH] ou [nomeH]    
    if sexo in 'Ff' and idade < 20:
        mulher20 += 1
    #Esse IF analisa e adiciona a variavel [mulher20] quando ela tem menos de 20 anos
    
mediaidade = somaidade / 4 

print('A média de idade do grupo é de {} anos!'.format(mediaidade))    
print('O homem mais velho tem {} e se chama {}!'.format(idadeH,nomeH))   
print('Ao todo são {} mulheres com menos de 20 anos! '.format(mulher20)) 

    
#Exercicio057: Faça um programa que leia o sexo de uma pessoa, mas só aceite os valores ‘M’ ou ‘F’. Caso esteja errado, peça a digitação novamente até ter um valor correto.
sexo = str(input('Qual o seu sexo [F/M]:')).strip().upper()[0]

while sexo not in 'MmFf':
    sexo = str(input('Dados Invalidos. Por favor, qual o seu sexo [F/M]:')).strip().upper()[0]

print('Sexo {} registrado com sucesso.'.format(sexo))


#Exercicio058: Melhore o jogo do DESAFIO 28 onde o computador vai “pensar” em um número entre 0 e 10. 
#Só que agora o jogador vai tentar adivinhar até acertar, mostrando no final quantos palpites foram necessários para vencer.

from random import randint

#Mensagem do computador
print('------ JOGO DA ADIVINHAÇÃO ------')
random = randint(0,10) #Randomizando os números
print('Acabei de pensar em um número entre 0 e 10')
print('Será que você consegue adivinhar qual foi?')
print('---------------------------------')

#Escolha do Jogador
palpite = int(input('Qual o seu palpite? '))
tentativas = 0

while palpite != random: 
    tentativas += 1 #Calculando o n° de tentativas ate acertar!
    if palpite > random:
        palpite = int(input('Ops, menos... Me de outro palpite: '))
    elif palpite < random:
        palpite = int(input('Ops, mais... Me de outro palpite: '))

print('Você venceu! Acertou o número com {} tentativas'.format(tentativas))        


#Exercicio059: Crie um programa que leia dois valores e mostre um menu na tela: [] soma [] multiplicar [] maior numero [] novos numeros [] finalizar
from time import sleep

print('------------- PROGRAMA DE CALCULO -------------')  
n1 = float(input('Informe um número: '))
n2 = float(input('Informe outro número: '))
sair = False

while sair == False:
    print(''' Selecione qual a opção desejada:
    [1] Somar os números;
    [2] Multiplicar os números;
    [3] Mostrar o maior valor;
    [4] Novos Números;
    [5] Finalizar o Programa''')
    escolha = int(input('       >>>> Informe sua escolha: '))

    if escolha == 5:
        sair = True
        print('\033[1;30;41mFinalizando...\033[m')
        sleep(1)
    else:
        if escolha == 1:
            soma = n1 + n2
            print('\n\033[1;30;41mA soma entre {:.0f} e {:.0f} resulta em: {:.0f} \033[m'.format(n1,n2,soma))
            print('=-'*20)
            sleep(1)
        if escolha == 2:
            multiplicacao = n1 * n2
            print('\n\033[1;30;41mA multiplicacao entre {:.0f} e {:.0f} resulta em: {:.0f} \033[m'.format(n1,n2,multiplicacao))
            print('=-'*20)
            sleep(1)
        if escolha == 3:
            if n1 > n2:
                print('\n\033[1;30;41mO maior número informado é {:.0f}\033[m'.format(n1))
                print('=-'*20)
                sleep(1)
            elif n2 > n1:
                print('\n\033[1;30;41mO maior número informado é {:.0f} \033[m'.format(n2))
                print('=-'*20)
                sleep(1)
            else:
                print('\n\033[1;30;41mOs números {:.0f} e {:.0f} são iguais. \033[m'.format(n1,n2))
                print('=-'*20)
                sleep(1)
        if escolha == 4:
            n1 = float(input('Informe o novo número: '))
            n2 = float(input('Informe o novo número: '))
            print('=-'*20)
            sleep(1)
        if escolha > 5:
            print('\n\033[1;30;41mOpção Invalida. Tente novamente! \033[m')
            print('=-'*20)
            sleep(1)


print('\033[1;30;41mPrograma finalizado! \033[m')        


#Exercicio060: Faça um programa que leia um número qualquer e mostre o seu fatorial
# EXEMPLO: 5! = 5 x 4 x 3 x 2 x 1 = 120

# 1ª JEITO DE RESOLVER:
from math import factorial

num = int(input('Digite um número para calcular seu fatorial: '))
fatorial = factorial(num)

print('O faltorial de {} é {}'.format(num,fatorial))


# 2ª JEITO DE RESOLVER:
num = int(input('Digite um número para calcular seu fatorial: '))
c = num
fatorial = 1

print('Calculando {}!: '.format(num), end='')

while c > 0:
    print('{}'.format(c), end='') # Para deixar de forma fatorial - EXEMPLO: 5  4  3  2  1 
    print(' x ' if c > 1 else ' = ', end='') # Formatando o calculo fatorial para: 5 x 4 x 3 x 2 x 1 =
    # Enquanto [c] for MAIOR que 1 (que é o ultimo numero no calculo fatorial) o WHILE vai por [x] quando [c] for igual [==] a 1, vai colocar [=]
    fatorial *= c
    c -= 1 # Atua decrescendo o input [NUM] - SEM essa formula, o print iria ficar 5 x 5 x 5 x 5 x 5 infinito

print('{}'.format(fatorial))


#Exercicio061: Refaça o DESAFIO 51, lendo o primeiro termo e a razão de uma PA, mostrando os 10 primeiros termos da progressão usando a estrutura while.
print('Vamos calcular Progressão Aritmetica')
print('-='*10)

primeiro = int(input('Insira um numero: '))
razao = int(input('Insira a progressão: '))
termo = primeiro
cont = 1 

while cont <= 10:
    print('{} -> '.format(termo),end='')
    termo += razao # Aqui atua dentro dos 10 termos, determinando quanto de acresimo entre um termo e outro - vai haver. EX: 3 -> 6 -> 9 [acrescimo de 3 em 3]
    cont += 1 # Atua acresentando no input, sem essa formula o print iria repetir infinito o numero do input. Com ele, vai acresentando até chegar em 10 termos (while cont <= 10)

print('FIM')


#Exercicio062: Melhore o DESAFIO 61, perguntando para o usuário se ele quer mostrar mais alguns termos. O programa encerrará quando ele disser que quer mostrar 0 termos.
print('Vamos calcular Progressão Aritmetica')
print('-='*10)

primeiro = int(input('Insira um numero: '))
razao = int(input('Insira a progressão: '))
termo = primeiro
cont = 1 
total = 0 #Total de termos que ele vai me mostrar (variavel para substituir o numero fixo [10] do exercicio 61)
mais = 10 #Simula a primeira incidencia da [PA]. Sempre no primeiro INPUT, vamos mostrar os 10 primeiros termos da [PA]

while mais != 0: #Determinando que o LOOP seja parado quando o usuario digitar [0]
    total = total + mais #Logo no inicio do programa, a variavel [Total] já passa a valer 10 - devido a soma com a variavel [mais]
    while cont <= total:
        print('{} -> '.format(termo),end='')
        termo += razao #Determina o quanto de acresimo entre um termo e outro
        cont += 1 
    print('Pausa...')
    mais = int(input('Quantos termos você quer mostrar a mais: ')) # Aqui, caso ele coloque mais termos, o loop se reinicia no (while cont <= total:) so que dessa vez,
    # com o historico da rodada passada
print('Progressão finalizada com {} termos mostrados.'.format(total))


#Exercicio063: Escreva um programa que leia um número N inteiro qualquer e mostre na tela os N primeiros elementos de uma Sequência de Fibonacci
print('-='*10)
print('SEQUENCIA DE FIBONACCI')

#Sequencia de FIBONACCI se baseia: Começa em ZERO e UM, e após isso, é a soma de seus sucessores;
#EXEMPLO: 0 -> 1 -> 1(0+1) -> 2(1+1) -> 3(2+1) -> 5(3+2) -> 8(3+5) E ASSIM SUCESSIVAMENTE

n = int(input('Quantos termos você quer mostrar: '))

#Como sabemos que em FIBONACCI, os dois primeiros termos sempre vão ser [0] e [1]:
t1 = 0
t2 = 1

print('{} -> {}'.format(t1,t2),end='') #Mostrando na tela os dois valores imutaveis em FIBONACCI

#Iniciando o LOOP dos termos selecionados no INPUT.
count = 3 #Como já temos os dois primeiros termos (0 e 1) iremos iniciar a contagem do WHILE em 3, que é aonde iremos necessitar de calculo!

while count <= n: #Enquanto o [Count] for menor ou igual ao INPUT [n]
    t3 = t1 + t2 #Faça o terceiro termo de fibonacci, que é a soma de t1 + t2
    print(' -> {}'.format(t3),end='') #Mostrando o terceiro termo na tela!
    #Agora iremos movimentar as variaveis, como todo termo seguinte ao dois primeiros, e a soma de seus antecessores, iremos mover uma casa a frente [t1], [t2] e [t3]
    t1 = t2 # É como um movimento de casa, [t1] passa a ocupar a casa do [t2] 
    t2 = t3 # E [t2] passa a ocupar a casa do t3(que ja foi calculado no inicio)
    # Com isso, matematicamente, todos os termos apos o termo [t2] vão ser considerados [t3] - eles não vão seguir uma logica, por que pra fazer o loop, tivemos que 
    # fazer essa motivamentação de casas (para não codar mais e mais)
    count += 1 #Somando o count inicial [3] para cada loop do WHILE (Então se eu colocar que quero ver 5 termos, ele vai me mostrar os 5 numeros seguintes ao 3)
print(' -> FIM')
print('-='*10)


#Exercicio064: Crie um programa que leia vários números inteiros pelo teclado. O programa só vai parar quando o usuário digitar o valor 999, que é a condição de parada. 
#No final, mostre quantos números foram digitados e qual foi a soma entre eles (desconsiderando o flag).

n = 0 #Variavel que possibilita o looping
count = 0 #Variavel de contagem de numeros informados
sum = 0 #Variavel de soma dos numeros informados

while n != 999: # Gambiarra apenas para gerar o laço
    n = int(input('Digite um numero [999 para parar]: '))
    if n != 999: #INPUT: enquanto o usuario não escrever 999 (para parar o comando)
        count += 1 #Conte quantos numeros ele colocou!
        sum += n #Some os numeros colocados
    elif n == 999: #INPUT: usuario colocou o comando para parar [999]
        n = 999 #Altere a variavel n para [999] e force o BREAK do laço

print('Você digitou {} números e a soma entre eles foi {}'.format(count,sum))
#Como tanto a variavel count quanto o sum, estão identados no primeiro IF, isso quer dizer que so vamos contar e somar os numeros que respeitarem [n != 999]


#Exercicio065: Crie um programa que leia vários números. No final da execução, mostre a média entre todos os valores e qual foi o maior e o menor valores lidos. 
#O programa deve perguntar ao usuário se ele quer ou não continuar a digitar valores.

resposta = 'S'
soma = quantidade = media = maiornum = menornum = 0
#Quando varias variaveis vão receber 0, podemos escrever em PYTHON desta forma!

while resposta in 'Ss':
    num = int(input('Digite um número: '))
    soma += num
    quantidade += 1
    if quantidade == 1: #Declarando um IF qualquer para estabelecer criterios
        maiornum = menornum = num #Declarando que o primeiro valor dado entrada no INPUT, vai ser considerado o [maiornum] e [menornum] - por que de fato é!
    else: #Neste ELSE venho com as condições de analise do sistema
        if num > maiornum: #se os proximos numeros na variavel [num] forem MAIOR que o primeiro numero colocado em [maiornum]
            maiornum = num #então altere o dado da variavel [maiornum] para o dado novo que foi introduzido na variavel [num] 
        elif num < menornum: #se os proximos numeros na variavel [num] forem MENOR que o primeiro numero colocado em [menornum]
            menornum = num #então altere o dado da variavel [menornum] para o dado novo que foi introduzido na variavel [num] 
    resposta = str(input('Quer continuar? [S/N]')).upper().strip()[0]

media = soma / quantidade
print('Você digitou {} números e a média foi {:.2f}'.format(quantidade,media))    
print('O maior valor digitado foi {} e o menor valor digitado foi {}'.format(maiornum,menornum))


#Exercicio066
cont = sum = 0

while True:
    num = int(input('Digite um valor [Digite 999 para parar]: '))
    if num == 999:
        break
    cont += 1 
    sum += num

print(f'Você digitou {cont} números e a soma entre eles foi {sum}')


#Exercicio067
while True:
    tabuada = int(input('Quer ver a tabuada de qual valor? '))
    if tabuada < 0:
        print('Programa de Tabuada Encerrado. Volte Sempre!')
        break
    for t in range(1,11):
        print(f'{tabuada} x {t:2} = {tabuada*t}')
    print('=-'*10)


#Exercicio068
from random import randint
print('=-'*15)
print('VAMOS JOGAR PAR OU ÍMPAR')
print('=-'*15)

contador_vitorias = 0 

while True:
    jogador = int(input('Diga um valor: '))
    maquina = randint(0,10)
    soma_resultado = maquina + jogador

    resposta = ' '
    while resposta not in 'PpIi':
        resposta = str(input('Par ou Ímpar [P/I]')).strip().upper()[0]

    print(f'Você jogou {jogador} e o computador jogou {maquina}. ',end='')

    if soma_resultado % 2 == 0:
        print(f'A soma dos valores deu {soma_resultado} e o valor é PAR')
        if resposta in 'Pp':
            print('Você GANHOU!')
            print('--'*15)
            contador_vitorias += 1
        elif resposta in 'Ii':
            print('Você PERDEU!')
            print('--'*15)
            break

    if soma_resultado % 2 == 1:
        print(f'A soma dos valores deu {soma_resultado} e o valor é ÍMPAR')
        if resposta in 'Ii':
            print('Você GANHOU!')
            print('--'*15)
            contador_vitorias += 1
        elif resposta in 'Pp':
            print('Você PERDEU!')
            print('--'*15)
            break      
       
print(f'GAME OVER! Você venceu {contador_vitorias} vezes')


#Exercicio069
print(' CADASTRE UMA PESSOA ')
print('-='*20)

total_pessoas = total_maiores = total_homens = total_mulheres_menos_20 = 0

while True:
    idade = int(input('Idade: '))
    sexo = ' '
    while sexo not in 'MF':
        sexo = str(input('Sexo [M/F]: ')).strip().upper()[0]

    total_pessoas += 1

    if idade >= 18:
        total_maiores += 1
    if sexo == 'M':
        total_homens += 1
    if sexo == 'F' and idade < 20:
        total_mulheres_menos_20 += 1

    continuar = ' '
    while continuar not in 'SN':    
        continuar = str(input('Deseja continuar? [S/N] ')).strip().upper()[0]
    if continuar == 'N':
        print('Encerrando o Programa!')
        break

print('-='*20)
print(f'Total de cadastros {total_pessoas}')
print(f'Total de pessoas com mais de 18 anos: {total_maiores} pessoas')
print(f'Ao todo temos {total_homens} Homens cadastrados!')
print(f'E temos {total_mulheres_menos_20} mulheres com menos de 20 anos')


#Exercicio070 Crie um programa que leia o nome e o preço de vários produtos. O programa deverá perguntar se o usuário vai continuar ou não. No final, mostre:
# A) qual é o total gasto na compra. 
# B) quantos produtos custam mais de R$1000.
# C) qual é o nome do produto mais barato.

total_gasto = acima_1k = menor_preço = cont =  0
mais_barato = ''

while True:
    produto = str(input('Insira o nome do Produto: ')).strip().upper()
    valor_produto = float(input('Informe o valor do Produto: R$ '))

    total_gasto += valor_produto

    cont += 1
    if cont == 1 or valor_produto < menor_preço:
        menor_preço = valor_produto
        mais_barato = produto

    if valor_produto > 1000:
        acima_1k += 1 
    
    decisao = ' '
    while decisao not in 'SN':
        decisao = str(input('Deseja continuar cadastrando produtos? [S/N]')).strip().upper()[0]
    if decisao == 'N':
        break


print(f'O Total gasto nesta compra foi de R${total_gasto}')
print(f'Ao todo, tempos {acima_1k} que custaram mais de R$1.000,00')
print(f'O produto mais barato comprado foi o {mais_barato} no valor de R${menor_preço}')


#Exercicio071 Crie um programa que simule o funcionamento de um caixa eletrônico. 
#No início, pergunte ao usuário qual será o valor a ser sacado (número inteiro) e o programa vai informar quantas cédulas de cada valor serão entregues. 
#OBS: considere que o caixa possui cédulas de R$50, R$20, R$10 e R$1

valor = int(input('Que valor deseja sacar? R$ '))
total_saque = valor
cédulas = 50 #Determinamos de ínicio que seja R$ 50
total_cédulas = 0 #Contagem de cédulas

while True:
    if total_saque >= cédulas: # Se o total for menor ou igual ao valor da cédula [R$50] faça isso:
        total_saque -= cédulas #Calculo de subtração: retire o valor das cédulas[R$50 de padrão] do total informado [Input]
        total_cédulas += 1 #Conte quantas cedulas foi retirada desse total [No caso, o padrão é R$50]
    else: #SE NÃO, faça isso: [Aqui, quando o valor do total ficar menor que o valor padrão da cedula R$50]
        if total_cédulas > 0: #Esse IF é apenas para print das cédulas necessarias
            print(f'Total de {total_cédulas} cédulas de R${cédulas}')
        if cédulas == 50: #Se o valor do total for menor, e as cédulas forem 50 
            cédulas = 20 #ABAIXE OS VALOR DA CÉDULA PARA 20
        elif cédulas == 20: #Se o valor do total for menor, e as cédulas forem 20
            cédulas = 10 #ABAIXE OS VALOR DA CÉDULA PARA 10
        elif cédulas == 10: #Se o valor do total for menor, e as cédulas forem 10
            cédulas = 1 #ABAIXE OS VALOR DA CÉDULA PARA 01
        total_cédulas = 0
        if total_saque == 0:
            break     

#Exercicio072: Crie um programa que tenha uma dupla totalmente preenchida com uma contagem por extenso, de zero até vinte. 
#Seu programa deverá ler um número pelo teclado (entre 0 e 20) e mostrá-lo por extenso.
        
numeros_por_extenso = (
    "zero", "um", "dois", "três", "quatro", "cinco",
    "seis", "sete", "oito", "nove", "dez",
    "onze", "doze", "treze", "quatorze", "quinze",
    "dezesseis", "dezessete", "dezoito", "dezenove", "vinte"
)

while True:
    pergunta = int(input('Me diga um número de 0 a 20: '))

    if 0 <= pergunta <= 20:
        print(f'Você digitou o número {numeros_por_extenso[pergunta]}')
        continuar = str(input('Deseja continuar? [S/N] ')).strip().upper()[0]
        if continuar == 'N':
            break
            print('Encerrando o Programa')
    else:
        print ('Número Invalido!')
        break     


#Exercicio073: Crie uma tupla preenchida com os 20 primeiros colocados da Tabela do Campeonato Brasileiro de Futebol, na ordem de colocação. Depois mostre:
#a) Os 5 primeiros times. b) Os últimos 4 colocados. c) Times em ordem alfabética. d) Em que posição está o time da Chapecoense.
    
brasileirao_2017 = (
    "Corinthians", "Palmeiras", "Santos", "Grêmio", 
    "Cruzeiro", "Flamengo", "Vasco da Gama", "Chapecoense",
    "Atlético Mineiro", "Botafogo", "Atlético Paranaense", "Bahia", 
    "São Paulo", "Fluminense", "Sport Recife", "EC Vitória", "Coritiba",
    "Avaí", "Ponte Preta","Atlético Goianiense")

#print(f'Lista dos times {brasileirao_2017}')
print(f'Os 05 primeiros colocados são {brasileirao_2017[0:6]}')
print('-='*30)
print(f'Os últimos 04 colocados são {brasileirao_2017[-4:]}')
print('-='*30)
print(f'Os times em ordem alfabetica {sorted(brasileirao_2017)}')
print('-='*30)
print(f'O time da chapecoense está na {brasileirao_2017.index("Chapecoense") + 1}ª posição')


#Exercicio074: Crie um programa que vai gerar cinco números aleatórios e colocar em uma tupla. 
#Depois disso, mostre a listagem de números gerados e também indique o menor e o maior valor que estão na tupla.

from random import randint

tupla = (randint(1,10), randint(1,10), randint(1,10), randint(1,10), randint(1,10))

print('Os valores sorteados foram: ')
for n in tupla:
    print(f' {tupla}',end='')

print(f'\nO maior valor sorteado foi {max(tupla)}')   
print(f'O maior valor sorteado foi {min(tupla)}')   


#Exercicio075: Desenvolva um programa que leia quatro valores pelo teclado e guarde-os em uma tupla. No final, mostre:
# A) Quantas vezes apareceu o valor 9. B) Em que posição foi digitado o primeiro valor 3. C) Quais foram os números pares.

numeros = (int(input('Digite um numero: ')),
           int(input('Digite outro numero: ')),
           int(input('Digite mais um numero: ')),
           int(input('Digite o ultimo numero: ')))

print(f'Você digitou os valores {numeros}')

print(f'O número 09 apareceu {numeros.count(9)} vezes.')

if 3 in numeros:
    print(f'A posição do primeiro valor 03 é {numeros.index(3)+1}')
else:
    print('O valor 03 não foi digitado!')


print(f'Os números pares são ', end='')
for n in numeros:
    if n % 2 ==0:
        print(n)



#Exercicio076: Crie um programa que tenha uma tupla única com nomes de produtos e seus respectivos preços, na sequência. 
#No final, mostre uma listagem de preços, organizando os dados em forma tabular.
        
print('-'*50)
print(f'{"LISTAGEM DE MATERIAL":^50}')
print('-'*50)

tupla = ('Lapis', 1.75,
         'Borracha', 2.50,
         'Caderno', 15.09,
         'Estojo', 25,
         'Taransferidor', 4.20,
         'Compasso', 9.99,
         'Mochilha', 120.32,
         'Canetas', 22.30,
         'Livro', 34.90)
#Cada palavra na tupla assume uma posição: 0ª Lapis - 1ª 1.75 - 2ª Borracha ... e assim vai

for pos in range (0,len(tupla)):
    if pos % 2 == 0: #Aqui separamos a posição, como todos os nomes dos itens estão em posição PAR, primeiro formatamos o nome, FILTRANDO com a clausula PAR
        print(f'{tupla[pos]:.<35}',end='')
    else: # Pegamos o resto para formatar, no caso os numero IMPAR (valores de cada categoria)
        print(f' R${tupla[pos]:>10.2f}')



#Exercicio077: Crie um programa que tenha uma tupla com várias palavras (não usar acentos). Depois disso, você deve mostrar, para cada palavra, quais são as suas vogais.
vogais = ('aprender','programar','linguagem',
          'python','curso','gratis','estudar',
          'praticar','trabalhar','mercado','programador','futuro')

for p in vogais:
    print(f'\nNa palavra {p.upper()} temos {"--------->":>5} ', end='')
    for letra in p:
        if letra.lower() in 'aeiou':
            print(f'{letra.upper()}', end='')      



#Exercicio078: aça um programa que leia 5 valores numéricos e guarde-os em uma lista. 
#No final, mostre qual foi o maior e o menor valor digitado e as suas respectivas posições na lista.
lista1 = []
maior = 0
menor = 0

for cont in range(0,5):
    lista1.append(int(input(f'Digite um valor para a posição {cont}:')))
    if cont == 0:
        maior = menor = lista1[cont]
    else:
        if lista1[cont] > maior:
            maior = lista1[cont]
        if lista1[cont] < menor:
            menor = lista1[cont]

print(f'O maior valor foi {maior} e suas posições: ', end='')
for i, v in enumerate(lista1):
    if v == maior:
        print(f' {i}...',end='')

print(f'\nO menor valor foi {menor} e suas posições: ', end='')
for i, v in enumerate(lista1):
    if v == menor:
        print(f' {i}...',end='')


#Exercicio079: Crie um programa onde o usuário possa digitar vários valores numéricos e cadastre-os em uma lista. 
#Caso o número já exista lá dentro, ele não será adicionado. No final, serão exibidos todos os valores únicos digitados, em ordem crescente.

lista2 = []

while True:
    numeros = (int(input('Digite um valor:')))
    
    if numeros not in lista2:
        lista2.append(numeros)
        print('Valor adicionado com sucesso...')
    else:
        print('Valor informado duplicado, não será acresentado')

    continuar = ' '
    while continuar not in 'SsNn':
        continuar = str(input('Deseja continuar? [S/N]')).strip().upper()[0]
    if continuar == 'N':
        break

lista2.sort #ORDERNANDO EM ORDEM CRESCENTE

print('-='*30)
print(f'Os valores informados foram: {lista2}')
        

#Exercicio080: Crie um programa onde o usuário possa digitar cinco valores numéricos e cadastre-os em uma lista, já na posição correta de inserção (sem usar o sort()). 
#No final, mostre a lista ordenada na tela.

lista3 = []

for c in range(0,5):
    n = int(input('Digite um valor:'))
    if c == 0 or n > lista3[-1]: #Se o valor inserido for o primeiro [OU] se o valor for maior que o ULTIMO VALOR (lista3[-1])
        lista3.append(n) #Adicione o valor a lista
    else:
        posicao = 0 #para descobrirmos as posições, iremos varrer o vetor inteiro, um a um
        while posicao < len(lista3): #Vai da posição ZERO até a ultima posicao da lista [LEN]
            if n <= lista3[posicao]: #Verifica se em cada posição da lista, o numero que quero inserir é menor ou igual a ele mesmo
                lista3.insert(posicao, n) # se for, ele vai inserir o valor[n], na posição [POS] que o laço está
                print(f'Adicionado na posição {pos} da lista')
                break
            posicao += 1 #Aqui serve apenas para o laço ir passando de posição em posição

print(f'Os valores digitados em ordem fica {lista3}')    


#Exercicio081: Crie um programa que vai ler vários números e colocar em uma lista. Depois disso, mostre:
#A) Quantos números foram digitados.  B) A lista de valores, ordenada de forma decrescente. C) Se o valor 5 foi digitado e está ou não na lista.

lista4 = []
contagem_numero = 0
contagem_cinco = 0

while True:
    n = int(input('Digite um valor:'))
    r = ' '

    lista4.append(n)
    contagem_numero += 1

    if n == 5:
        contagem_cinco += 1

    while r not in 'SsNn':
        r = str(input('Deseja Continuar: [S/N]')).strip().upper()[0]
    if r == 'N':
        break
   

lista4.sort(reverse=True)

print('-='*30)
print(f'A lista ficou: {lista4};')
print(f'Foram digitados {contagem_numero} valores ao todo;')

if contagem_cinco > 0:
    print(f'O valor 05 foi digitado {contagem_cinco} vezes. E consta na lista;')
else:
    print('O valor 05 não foi digitado;')    


#Exercicio082: Crie um programa que vai ler vários números e colocar em uma lista. 
#Depois disso, crie duas listas extras que vão conter apenas os valores pares e os valores ímpares digitados, respectivamente. 
#Ao final, mostre o conteúdo das três listas geradas.

lista5 = []
pares = []
impares = []
contagem_par = contagem_impar = 0

while True:
    n = int(input('Digite um valor:'))
    r = ' '

    lista5.append(n)

    if n % 2 == 0:
        pares.append(n)
        contagem_par += 1
    if n % 2 == 1:
        impares.append(n)
        contagem_impar += 1
   
    while r not in 'SsNn':
        r = str(input('Deseja Continuar: [S/N]')).strip().upper()[0]
    if r == 'N':
        break

pares.sort
impares.sort

print('-'*30) 
print(f'LISTA COM TODOS OS VALORES: {lista5}')   
print(f'LISTA COM OS {contagem_par} VALORES PARES DIGITADOS:{pares}') 
print(f'LISTA COM OS {contagem_impar} VALORES PARES DIGITADOS:{impares}') 
    

#Exercicio083: Crie um programa onde o usuário digite uma expressão qualquer que use parênteses. 
#Seu aplicativo deverá analisar se a expressão passada está com os parênteses abertos e fechados na ordem correta.
expressao = str(input('Digite uma expressão'))
lista6 = []

for simbolo in expressao:
    if simbolo == '(': #Toda vez que abrir um parenteses, vai ser adicionado na lista6 um parenteses
        lista6.append('(') #adicão do '('
    elif simbolo == ')': #E se for o fechamento de um parenteses ')'
        if len(lista6) > 0: #Bem, se a tamanho da lista for maior que 0 [>0} (ou seja, se tiver dados, sendo que so é colocado dados com '(')
            lista6.pop() # Essa função exclui o ultimo dado da lista. (NOTE QUE NAO ADICIONEI AQUI NENHUM PARENTESES A MAIS NA LISTA, APENAS EXCLUI O QUE JA HAVIA)
        else: #Se o tamanho da lista nao for maior que zero (ou seja, nao tem dados na lista)
            lista6.append(')') #AI SIM, adicione um ')' na lista (NOTE QUE AQUI ESTÁ SOBRANDO PARENTESES)
            break

#basicamente nesse for, estaremos juntando os pares, enquanto a lista for maior que 0 [> 0] ou seja, tiver dados de '(' juntamos os pares com ')'
#Se não tiver nenhum dado, ele adiciona e fica sobrando (sozinho) algum dos dois parenteses

if len(lista6) == 0:
    print('Sua expressão está correta!')
else:
    print('Sua expressão está errada!')                



#Exercicio084: Faça um programa que leia nome e peso de várias pessoas, guardando tudo em uma lista. No final, mostre:
# A) Quantas pessoas foram cadastradas. B) Uma listagem com as pessoas mais pesadas. C) Uma listagem com as pessoas mais leves.
lista7 = [] #Lista principal
temporaria = [] #Lista auxiliar
maior_peso = menor_peso = 0

while True:
    temporaria.append(str(input('Digite o nome: '))) #Cadastro na lista temporaria [0]
    temporaria.append(float(input('Digite o peso: '))) #Cadastro na lista temporaria [1]

    #Analistando as variaveis maior_peso = menor_peso
    if len(lista7) == 0: #Se o tamanho(LEN) da lista7(PRINCIPAL) for igual a 0. Ou seja, se não tiver nenhum dado na lista principal.
        maior_peso = menor_peso = temporaria[1] #Determine que o primeiro dado colocado na lista, seja considerado o maior_peso e menor_peso
    else: #Se não for o primeiro dado da lista7 (principal)
        if temporaria[1] > maior_peso: #Se o valor inserido na lista temporaria for maior que o valor de maior_peso registrado
            maior_peso = temporaria[1] # Altere o maior_peso pro valor recem cadastrado em temporaria
        elif temporaria[1] < menor_peso: #Se o valor inserido na lista temporaria for menor que o valor de menor_peso registrado
            menor_peso = temporaria[1] # Altere o menor_peso pro valor recem cadastrado em temporaria

    #Note que coloco temporaria[1]; o significado desse [1] e a localidade do dado na lista, [0] seria a localização dos nomes e [1] os da idade
    #Como é uma lista dentro de uma lista, temos a localização do conjunto de dados (Nome e Idade) e dentro desse conjunto temos a localização distinta dos dados
    # No caso (Nome[0] e Peso[1]) ---> Isso, também pode estar na localidade [0] na primeira instancia da lista (como é uma lista, dentro de outra lista)                       

    lista7.append(temporaria[:]) #Criando uma cópia dos dados da lista temporaria, para a lista principal - Utilizando o [:]
    temporaria.clear() #Assim que criado a cópia entre as duas, limpamos os dados da lista temporaria

    r = str(input('Deseja Continuar: [S/N]')).strip().upper()[0] #Clausula para interromper o laço
    if r == 'N':
        break

print('-='*30)
print(f'Ao todo foram cadastradas {len(lista7)} pessoas.') 
#Aqui em vez de utilizar um contador para o QTD de pessoas, utilizei o tamanho da lista7, que vai retornar o n° de cadastros

print(f'O maior peso foi de {maior_peso}Kg. Peso de ', end='')
for peso in lista7: #Varremos com o FOR a lista7 para localizar os nomes das pessoas com maior peso
    if peso[1] == maior_peso: #Clausula para localizar --- lembrando que o [1] é a localidade do dado PESO na lista
        print(f'[{peso[0]}]',end='') #Retorna um print com o nome da pessoa com o maior peso -- lembrando que o [0] é a localidade do dado NOME na lista

print()

print(f'O menor peso foi de {menor_peso}Kg. Peso de ', end='')
for peso in lista7: #Varremos com o FOR a lista7 para localizar os nomes das pessoas com menor peso
    if peso[1] == menor_peso: #Clausula para localizar --- lembrando que o [1] é a localidade do dado PESO na lista
        print(f'[{peso[0]}]',end='') #Retorna um print com o nome da pessoa com o MENOR peso -- lembrando que o [0] é a localidade do dado NOME na lista   



#Exercicio085: Crie um programa onde o usuário possa digitar sete valores numéricos e cadastre-os em uma lista única que mantenha separados os valores pares e ímpares. 
#No final, mostre os valores pares e ímpares em ordem crescente.
lista8 = [[], []] #Lista separada por: [ [0 PAR] , [1 IMPAR] ]
valor = 0

for numeros in range(1,8):
    valor = (int(input(f'Digite o {numeros}° valor: ')))

    if valor % 2 == 0: #Se for par, coloca na lista no localidade [0]
        lista8[0].append(valor)
    elif valor % 2 == 1: #Se for par, coloca na lista no localidade [1]
        lista8[1].append(valor)

lista8[0].sort() #Ordenando em ordem crescente
lista8[1].sort() #Ordenando em ordem crescente

print(f'os valores pares digitados foram: {lista8[0]}')
print(f'os valores ímpares digitados foram: {lista8[1]}')


#Exercicio086: Crie um programa que declare uma matriz de dimensão 3×3 e preencha com valores lidos pelo teclado. No final, mostre a matriz na tela, com a formatação correta.
matriz = [[0, 0, 0],[0, 0, 0],[0, 0, 0]]

for linha in range (0,3): #Laço FOR para alimentação - Aqui so estamos colocando os valores dentro da matriz
    for coluna in range(0,3):
        matriz[linha][coluna] = (int(input(f'Digite um valor para [{linha}, {coluna}]: ')))

for linha in range (0,3): #Esses são FOR de estrutura de repetição, para mostrar os dados na tela em forma 3x3
    for coluna in range(0,3):
        print(f'[{matriz[linha][coluna]:^5}]', end='') #Aqui ele vai mostrar os dados um do lado do outro
    print() #VOLTEI UMA IDENTAÇÃO, e fiz a quebra de linha para cada incidencia de 03 no range do FOR, ai forma 3x3       


#Exercicio087: Aprimore o desafio anterior, mostrando no final:                                                    
#A) A soma de todos os valores pares digitados. B) A soma dos valores da terceira coluna. C) O maior valor da segunda linha.
matriz = [[0, 0, 0],[0, 0, 0],[0, 0, 0]]
soma_pares = maior_valor = terceira_coluna = 0

for linha in range (0,3): 
    for coluna in range(0,3):
        matriz[linha][coluna] = (int(input(f'Digite um valor para [{linha}, {coluna}]: ')))

print('-='*30)

for linha in range (0,3): 
    for coluna in range(0,3):
        print(f'[{matriz[linha][coluna]:^5}]', end='')
        if matriz[linha][coluna] % 2 == 0:
            soma_pares += matriz[linha][coluna]
    print()  

print('-='*30)
print(f'A soma dos valores pares é {soma_pares}')

for linha in range(0,3): #Fazendo um FOR apenas para a linha, por que a coluna é fixa (coluna 02 para o python, e coluna tres para gente)
    terceira_coluna += matriz[linha][2] #Soma a variavel [terceira_coluna] com os numeros constantes na (matriz[linha][2]) - Lembrando que linha varia, somente a coluna que é fixa
    #No caso aqui fica [2] por que para o python, é indexado em 0,1,2... então pra gente a coluna 3 para o python é a coluna 2
print(f'A soma dos valores da terceira coluna é {terceira_coluna}')

for coluna in range(0,3): #Mesma premissa que o FOR anterior, aqui AS LINHAS SÃO FIXA
    if coluna == 0: #Se for o primeiro dado entrando na variavel COLUNA do range
        maior_valor = matriz[1][coluna] #determine como o maior valor
    elif matriz[1][coluna] > maior_valor: #se algum valor colocado no input for maior que o maior_valor, altere o valor para o novo valor
        maior_valor = matriz[1][coluna]
    #LEMBRANDO QUE AQUI ESTÁ FIXADO APENAS A ANALISE NA LINHA 2 (para o python é linha [1] por que é indexado em 0)    
print(f'O maior valor da segunda linha é {maior_valor}')      


#Exercicio088:Faça um programa que ajude um jogador da MEGA SENA a criar palpites.
#O programa vai perguntar quantos jogos serão gerados e vai sortear 6 números entre 1 e 60 para cada jogo, cadastrando tudo em uma lista composta.
from random import randint
from time import sleep

lista10 = list()
lista_copia = list()

print('-'*30)
print('      JOGO DA MEGA SENA      ')
print('-'*30)

jogada = int(input('Quantos jogos você deseja fazer? '))
print('-'*30)
total_jogada = 1

while total_jogada <= jogada:
    contador = 0 
    while True:
        num = randint(1,60)
        if num not in lista10:
            lista10.append(num)
            contador += 1 
        if contador >= 6:
            break
    lista10.sort()
    lista_copia.append(lista10[:])
    lista10.clear()
    total_jogada += 1 

print('-='*3, f'Sorteando {jogada} Jogos', '-='*3)
for i, l in enumerate(lista_copia):
    print(f'Jogo {i+1}: {l}')
    sleep(1)

print('-='*3, 'Boa Sorte!!', '-='*3) 


#Exercicio089: Crie um programa que leia nome e duas notas de vários alunos e guarde tudo em uma lista composta. 
#No final, mostre um boletim contendo a média de cada um e permita que o usuário possa mostrar as notas de cada aluno individualmente.

ficha = []

while True:
    nome = str(input('Nome: '))
    nota1 = float(input('Nota 1: '))
    nota2= float(input('nota 2: '))
    media = (nota1 + nota2) / 2

    ficha.append([nome, [nota1, nota2], media])
    #Na lista, os dados vão ser [Nome(0), [nota1, nota2](1), Média(2)]

    resp = str(input('Deseja continuar? [S/N] '))
    if resp in 'Nn':
        break
    #esse While retorna: [['Lucas', [10.0, 8.4], 9.2], ['Bia', [6.0, 8.5], 7.25]]
    #Ou seja, ele retorna uma lista dentro de uma lista, com os dados.. Então temos [Nome, [Nota 1, Nota 2], Média]

print('-='*30)
print(f'{"N°":<4}{"Nome":<10}{"Média":>8}')
print('-'*26)

for indice, aluno in enumerate(ficha):
    print(f'{indice:<4}{aluno[0]:<10}{aluno[2]:>8.1f}')
    #Nesse for, trazemos os dados dos alunos em um PRINT. Neste formato: [Nome(0), [nota1, nota2](1), Média(2)]
    #O indice retrata em qual posição está o dado referido.
    # aluno[0]: Faz menção a localização do dado na lista FICHA. neste caso, ele está puxando o NOME[0]
    # aluno[2]: Faz menção a MÉDIA da nota deste aluno; No caso da [nota1, nota2], como é uma lista, eles assumem somente uma posição dentro da lista geral [1]
    # {:<4}{:<10}{:>8.1f} Tudo isso é apenas formatação para sair de forma tabular os dados

while True:
    print('-'*35) 
    opc = int(input('Mostrar notas de qual aluno? [999 interrompe]: '))
    if opc == 999:
        print('Programa Finalizando')
        break

    if opc <= len(ficha)-1:
        print(f'Notas de {ficha[opc][0]} são {ficha[opc][1]}')


#Exercicio090: Faça um programa que leia nome e média de um aluno, guardando também a situação em um dicionário. No final, mostre o conteúdo da estrutura na tela.
classe = {}
# ou classe = dict()

classe['Alunos'] = str(input('Digite o aluno: '))
classe['Média'] = float(input(f'Digite a média do aluno {classe["Alunos"]}: '))

if classe["Média"] < 5:
    classe['Situação'] = 'Reprovado'
elif classe["Média"] > 7:
    classe['Situação'] = 'Aprovado'
else:
    classe['Situação'] = 'Em recuperação'

print('-='*30)
print(f'O {classe["Alunos"]} foi {classe["Situação"]} com uma média de {classe["Média"]}')


#Exercicio091: Crie um programa onde 4 jogadores joguem um dado e tenham resultados aleatórios. Guarde esses resultados em um dicionário em Python. 
#No final, coloque esse dicionário em ordem, sabendo que o vencedor tirou o maior número no dado.
from random import randint
from time import sleep
from operator import itemgetter

jogadas = {'Breno': randint(1,6),
           'Joaquim': randint(1,6),
           'Alex': randint(1,6),
           'Lucas': randint(1,6)}
#{Jogador POSIÇÃO(0) : Dado POSIÇÃO(1)}

rank = []
print('SORTEANDO NÚMEROS:')
sleep(1)

for i, j in jogadas.items():
    print(f'O dado sorteou {j}!')
    sleep(0.8)

rank = sorted(jogadas.items(), key=itemgetter(1), reverse=True)
#Aqui ele ordena o rank com base no resultado do [randint(1,6)] - Na query, o que faz menção a esse resultado é [key=itemgetter(1)]
#Ressaltando que esse resultado gera uma LISTA. Logo não tem como utilizar .items()

print('-'*30)
print('Ordenando jogadores...')
sleep(0.8)
print(f'\n{"POSIÇÃO":^10} {"JOGADOR":<10} {"SORTEIO":^10}')
for i, j in enumerate(rank):
    print(f'{i+1:^10} {j[0]:<10} {j[1]:^10}')
    sleep(1)
print('-'*30)


#Exercicio092: Crie um programa que leia nome, ano de nascimento e carteira de trabalho e cadastre-o (com idade) em um dicionário. Se por acaso a CTPS for diferente de ZERO, 
#o dicionário receberá também o ano de contratação e o salário. Calcule e acrescente, além da idade, com quantos anos a pessoa vai se aposentar.
from datetime import date
from time import sleep

anoatual = date.today().year
cad = {}

cad['Nome'] = str(input('Digite seu nome: '))
ano_nasc = int(input('Digite seu ano de nascimento: '))
cad['Idade'] = anoatual - ano_nasc
cad['CTPS'] = int(input('Digite o Nº da carteira[0 se não tiver]: '))

if cad['CTPS'] != 0:
    cad['Contratação'] = int(input('Informe o ano de contratação: '))
    cad['Salario'] = float(input('Informe o salario R$ '))
    cad['Aposentadoria'] = cad['Idade'] + ((cad['Contratação'] + 35) - anoatual)

    print('\nPROCESSANDO DADOS!')
    sleep(1)
    print('-'*80)
    print(f'{"Nome":<15} {"Idade":<7} {"CTPS":<10} {"Contratação":<15} {"Salario":<10} {"Aposentadoria":<15}')
    print(f'{cad["Nome"]:<15} {cad["Idade"]:<7} {cad["CTPS"]:<10} {cad["Contratação"]:<15} {cad["Salario"]:<10} {cad["Aposentadoria"]:<15}')
    print('-'*80)

elif cad['CTPS'] == 0:
    print('\nPROCESSANDO DADOS!')
    sleep(1)

    print('-'*40)
    print(f'{"Nome":<15} {"Idade":<7} {"CTPS":<10}')
    print(f'\n{cad["Nome"]:<15} {cad["Idade"]:<7} {"NÃO APLICA":<10}')
    print('-'*40)


#Exercicio093: Crie um programa que gerencie o aproveitamento de um jogador de futebol. O programa vai ler o nome do jogador e quantas partidas ele jogou. 
#Depois vai ler a quantidade de gols feitos em cada partida. No final, tudo isso será guardado em um dicionário, incluindo o total de gols feitos durante o campeonato.

jogador = {}
list_gol = []

jogador['Nome'] = str(input('Digite o nome do jogador: '))
jogador['Nº de Partidas'] = int(input('Informe quantas partidas ele jogou: '))

for n_partidas in range(0,jogador['Nº de Partidas']):
    list_gol.append(int(input(f'Informe quantos gols ele fez no {n_partidas+1}º jogo: ')))

jogador['Saldo de Gols'] = list_gol[:]
jogador['Total de Gols'] = sum(list_gol)

print('--'*30)
print(f'{"Partida":<15} {"Jogador":<15} {"Saldo de Gols":<30}')
for i, j in enumerate(jogador['Saldo de Gols']):
    print(f'{i+1:<15} {jogador["Nome"]:<15} {j:<30}')
print('--'*30)

print(f'\033[1;30;41mO jogador {jogador["Nome"]} participou de {jogador["Nº de Partidas"]} jogos e fez um total de {jogador["Total de Gols"]} gols!\033[m')
print(jogador)


#Exercicio094: Crie um programa que leia nome, sexo e idade de várias pessoas, guardando os dados de cada pessoa em um dicionário e todos os dicionários em uma lista. 
#No final, mostre: A) Quantas pessoas foram cadastradas B) A média de idade C) Uma lista com as mulheres D) Uma lista de pessoas com idade acima da média


geral_cadastro = []
cadastro = {}
soma_idade = media_idade = 0

while True:
    cadastro.clear()
    cadastro['Nome'] = str(input('Digite o nome: '))

    cadastro['Idade'] = int(input('Digite a idade: '))
    soma_idade += cadastro['Idade']

    cadastro['Sexo'] = str(input('Digite o sexo [M/F]: ')).upper().strip()[0]
    while cadastro['Sexo'] not in 'MmFf':
        cadastro['Sexo'] = str(input('ERRO, FAVOR INFORMAR - Digite o sexo [M/F]: ')).upper().strip()[0]
    
    geral_cadastro.append(cadastro.copy())

    continuar = str(input('Deseja continuar? [S/N]: ')).upper().strip()[0]
    while continuar not in 'SsNn':
        continuar = str(input('ERRO, FAVOR INFORMAR - Deseja continuar? [S/N]: ')).upper().strip()[0]
    if continuar == 'N':
        print(f'\033[1;30;41mEncerrando Programa!\033[m')
        break

media_idade = soma_idade / len(geral_cadastro)

print('-='*45)
print(f'Ao todo, foram cadastrados {len(geral_cadastro)} pessoas!') #A) Quantas pessoas foram cadastradas

print(f'A média de idade dos cadastros é de {media_idade:.0f} anos!') #B) A média de idade

print('Os cadastros femininos foram: ', end='') #C) Uma lista com as mulheres
for pessoas in geral_cadastro: 
    if pessoas['Sexo'] == 'F':
        print(f'{pessoas["Nome"]} | ', end='') 

print()

print('Os usuarios com a idade acima da media foram: ', end='') #D) Uma lista de pessoas com idade acima da média
for pessoas in geral_cadastro: 
    if cadastro['Idade'] > media_idade:
        print(f'{pessoas["Nome"]} | ', end='') 

        
#Exercicio095: Aprimore o desafio 93 para que ele funcione com vários jogadores, incluindo um sistema de visualização de detalhes do aproveitamento de cada jogador.

jogador = {}
list_gol = []
time = []

while True:
    jogador.clear()
    jogador['Nome'] = str(input('Digite o nome do jogador: '))
    jogador['Nº de Partidas'] = int(input('Informe quantas partidas ele jogou: '))
    list_gol.clear()

    for n_partidas in range(0,jogador['Nº de Partidas']):
        list_gol.append(int(input(f'Informe quantos gols ele fez no {n_partidas+1}º jogo: ')))

    jogador['Saldo de Gols'] = list_gol[:]
    jogador['Total de Gols'] = sum(list_gol)
    time.append(jogador.copy())

    continuar = str(input('Deseja continuar? [S/N]: ')).upper().strip()[0]
    while continuar not in 'SsNn':
        continuar = str(input('ERRO, FAVOR INFORMAR - Deseja continuar? [S/N]: ')).upper().strip()[0]
    if continuar == 'N':
        print(f'\033[1;30;41mEncerrando Programa!\033[m')
        break 

print(f'{"Cód ":<3} {"Jogador":<15} {"Nº Partidas":<15} {"Saldo de Gols":<15} {"Total de gols":<15}')   
print('--'*35)
for k, v in enumerate(time):
    print(f'{k:<3} ',end='')
    for d in v.values():
        print(f' {str(d):<15} ',end='')
    print()        
print('--'*35)

while True:
    busca = int(input('Qual jogador deseja verificar? [999 PARAR] '))
    if busca == 999:
        break
    if busca >= len(time):
        print(f'ERRO! Não existe o jogador com código {busca}')
    else:
        print(f'\033[1;30;41mLevantamento do jogador {time[busca]["Nome"]}\033[m')
        for i, g in enumerate(time[busca]["Saldo de Gols"]):
            print(f'No jogo {i+1} fez {g} gols.')

            
#Exercicio096: Faça um programa que tenha uma função chamada área(), que receba as dimensões de um terreno retangular (largura e comprimento) e mostre a área do terreno.

def titulo(msg):
    print('―'*len(msg))
    print(msg)
    print('―'*len(msg))


def area(largura,comprimento):
    area = l * c
    print(f'A area de um terreno {largura} x {comprimento} é de {area}m²')


titulo('Controle de Terrenos')
l = float(input('Informe a Largura (M):'))
c = float(input('Informe o Comprimento (M):'))
area(l,c)  


#Exercicio097: Faça um programa que tenha uma função chamada escreva(), que receba um texto qualquer como parâmetro e mostre uma mensagem com tamanho adaptável.
def titulo(msg):
    print('―'*(len(msg)+2))
    print(f' {msg}')
    print('―'*(len(msg)+2))

titulo('Lucas')
titulo('Curso em video de Python')
titulo('CeV')



#Exercicio098: Faça um programa que tenha uma função chamada contador(), que receba três parâmetros: início, fim e passo. 
#Seu programa tem que realizar três contagens através da função criada: a) de 1 até 10, de 1 em 1 - b) de 10 até 0, de 2 em 2 - c) uma contagem personalizada 
from time import sleep

def contador(inicio,fim,passo):

    if passo < 0: #Se o contador for em ordem DESCRESENTE, mudamos o sinal subtraindo por -1. Isso vai evitar problema, pois se não ele vai infitinamente calcular
        passo *= -1
    elif passo == 0: #Se colocar o passo em 0, ele automaticamente muda para os numeros passarem de 1 em 1.
        passo = 1

    print(f'Contagem de {inicio} até {fim} de {passo} em {passo}:')
    
    if inicio < fim: #Se a contagem for crescente (EX: 01, 02, 03, 04, 05...)
        contador = inicio
        while contador <= fim:
            print(f' {contador} ->', end='', flush=True) #flush=True PARA REMOVER O BUFFER DE TELA ea função sleep funcionar
            sleep(0.5)
            contador += passo
        print(' Fim!')

    else: #Se a contagem for decrescer (EX: 10, 09, 08, 07, 06...)
        contador = inicio
        while contador >= fim:
            print(f' {contador} ->', end='', flush=True) #flush=True PARA REMOVER O BUFFER DE TELA ea função sleep funcionar
            sleep(0.5)
            contador -= passo
        print(' Fim!')
    
# a) de 1 até 10, de 1 em 1:   
contador(0,10,1)
print('-='*30) 

#b) de 10 até 0, de 2 em 2:   
contador(10,0,2)
print('-='*30)

#c) uma contagem personalizada
print('Escolha a contagem:')
i = int(input('Ínicio: '))
f = int(input('Fim:    '))
p = int(input('Passo:  '))
print('-='*30)
contador(i,f,p)


#Exercicio099: Faça um programa que tenha uma função chamada maior(), que receba vários parâmetros com valores inteiros. 
#Seu programa tem que analisar todos os valores e dizer qual deles é o maior.
from time import sleep
from random import randint

def maior(*num):
    contador = maior_valor = 0
    print('Analisando os valores passados:')

    for valor in num:
        print(f'{valor} -> ',end='', flush=True)
        sleep(0.3)
        if contador == 0:
            maior_valor = valor
        else:
            if valor > maior_valor:
                maior_valor = valor
        contador += 1
    print('Fim!',end='')        
    print(f'\nForam informados {contador} valores.', end='')
    print(f' O maior valor informado foi {maior_valor}!')
    print('-='*30)

maior(4,5,6,8,9,10,2,3,5,4,65)
maior(10,20,30,55)
maior(99,100)


#Exercicio100: Faça um programa que tenha uma lista chamada números e duas funções chamadas sorteia() e somaPar(). 
#A primeira função vai sortear 5 números e vai colocá-los dentro da lista e a segunda função vai mostrar a soma entre todos os valores pares sorteados pela função anterior.
from random import randint
from time import sleep

numeros = list()

def sorteia(lista):
    print('EFETUANDO O SORTEIO DE VALORES:')
    for contador in range(0,5):
        n = randint(1,10)
        lista.append(n)
        print(f'{n} -> ', end='', flush=True)
        sleep(0.3)
    print('Fim!')  
    print('-='*30)

def somapar(lista):
    soma = 0
    for valor in lista:
        if valor % 2 == 0:
            soma += valor
    print(f'Somando os valores pares da {lista}, temos a soma de {soma}!') 
    print()       


sorteia(numeros)
somapar(numeros)

#Exercicio101: Crie um programa que tenha uma função chamada voto() que vai receber como parâmetro o ano de nascimento de uma pessoa. 
#retornando um valor literal indicando se uma pessoa tem voto NEGADO, OPCIONAL e OBRIGATÓRIO nas eleições.
def voto(ano):
    from datetime import date 
    #Importei a biblioteca dentro da função - ISSO ECONOMIZA MEMORIA e utilizo a biblioteca SOMENTE DENTRO DA FUNÇÃO

    anoatual = date.today().year
    idade = anoatual - ano

    if idade > 18 and idade < 65:
        return f'Você possui {idade} anos. O VOTO É OBRIGATORIO!'
    elif idade < 16:
         return f'Você possui {idade} anos. AINDA NÃO VOTA!'
    elif idade > 65 or 16 <= idade < 18:
        return f'Você possui {idade} anos. O VOTO É OPCIONAL!'


pergunta = int(input('Em que ano você nasceu? '))
print(voto(pergunta))


#Exercicio102:Crie um programa que tenha uma função fatorial() que receba dois parâmetros: o primeiro que indique o número a calcular
# e outro chamado show, que será um valor lógico (opcional) indicando se será mostrado ou não na tela o processo de cálculo do fatorial.

def fatorial(num, show=False):
    """
    -> Calcula o fatorial de um número.
    :param n -> O número a ser calculado
    :param show -> (opcional) mostrar ou não a conta
    :return -> O valor do fatorial do número N
    """
    
    fatorial = 1

    for contador in range (num,0,-1): # calculo fatorial vai decresendo: EX - 5 X 4 X 3 X 2 X 1
        if show: #Apenas para caso deseje apresentar a formula de calculo
            print(contador, end='') # adiciona os numeros do in range (num,0,-1)
            if contador > 1: #Coloca um x entre os numeros (5 x 4 x 3...)
                print(' x ', end ='')
            else: #Como o in range está descrecendo, enquanto o contador for maior que 1(coloca x) quando contador não for maior (ou seja 0) quer dizer que chegou no fim do laço;
                print(' = ', end='') #Ai coloca um = para definir como final do laço
        fatorial *= contador #a cada num no range, ele multiplica o fatorial pelo num no range
    return fatorial     


print(fatorial(5,show=True))
help(fatorial)   


#Exercicio103: Faça um programa que tenha uma função chamada ficha(), que receba dois parâmetros opcionais: o nome de um jogador e quantos gols ele marcou.
# O programa deverá ser capaz de mostrar a ficha do jogador, mesmo que algum dado não tenha sido informado corretamente.

def ficha(jogador='<não informado>', gol=0):
    print(f'O jogador {jogador} fez {gols} gol(s) no campeonato;')
#Dentro da DEF, quando determinamos os parametros, eles se tornam opcionais.
#Neste caso, jogador='Não informado' por padrão é NÃO INFORMADO, essa clasula so é alterada se no input informarmos o nome
#Mesma coisa vale pros gols!    


#Programa principal:
player = str(input('Digite o nome do jogador: '))
gols = str(input(f'Saldo de gols do jogador {player}: '))
#Deixamos o saldo de gols em STR por que apenas STRINGS permitem que o input avançe mesmo estando em branco (EM INT, ele da erro no codigo)

if gols.isnumeric(): #Validando a informação do SALDO DE GOLS - SE CONSTAR INFORMAÇÃO E ELA FOR NUMERICA (is numeric)
    gols = int(gols) #Transformamos o tipo da variavel gols para inteiro (ou seja de STR passa a ser INT)
else: #Se não tiver informações NUMERICAS
    gols = 0 #Muda o valor da variavel gols para 0, independente do que esteja escrito!

if player.strip() == '': #Validando as informações da variavel PLAYER. para retornar conforme situações;
    ficha(gol = gols) #Se a variavel player estiver VAZIA, retorne apenas a DEF FICHA com o saldo de gols informado
else: #se a variavel NÃO ESTIVER VAZIA, retorne a DEF FICHA com os dados informados no INPUT
    ficha(player, gols)  


#Exercicio104:Crie um programa que tenha a função leiaInt(), 
#que vai funcionar de forma semelhante ‘a função input() do Python, só que fazendo a validação para aceitar apenas um valor numérico.
def leiaInt(msg):
    ok = False
    valor = 0

    while True:
        n = str(input(msg)) #TRANSFORMA A VARIAVEL N (que é o programa princial) em um INPUT
        if n.isnumeric(): # Se N for um valor numerico
            valor = int(n) #Valor recebe o valor de N(inteiro)
            ok = True #Altera a variavel OK para TRUE
        else: #Se o valor informado na variavel N não for NUMERICO
            print('\033[0;31mERRO! Digite um número válido!\033[m') # da mensagem de erro e o laço fica se repetindo até ser numerico
        if ok: #AQUI BASICAMENTE É SE O OK = TRUE --- Quebre o laço
            break
    return valor #Retorna o valor informado após validado as informaçoes no laço

#Programa Principal:
n = leiaInt('Digite um número: ')
print(f'Você acabou de digitar o número {n}')


#Exercicio105: Faça um programa que tenha uma função notas() que pode receber várias notas de alunos e vai retornar um dicionário com as seguintes informações:
# – Quantidade de notas – A maior nota – A menor nota  – A média da turma  – A situação (opcional)
def notas(*n, sit=False):
    """
    -> Função para analisar notas e situações de vários alunos;
    :param n: uma ou mais notas dos alunos (aceita varias)
    :param sit: (VALOR OPCIONAL), indicando se deve ou não adicionar a situação do aluno;
    :return : dicionario com várias informações sobre a situação;
    """
    r = dict()
    r['total'] = len(n)
    r['média'] = sum(n) / len(n)
    r['maior_elemento'] = max(n)
    r['menor_elemento'] = min(n)
    #Com base na chamada da função notas(); todo valor incluido dentro de (), vai ser introduzido no dicionario r;
    #Esse dicionario, com os valores vai trazer, total de numeros colocados, media entre eles, maior e menor
    #Basicamente esse é o programa!

    if sit: #Quando é um valor booleano, não se faz necessario declarar sit=True;
        if r['média'] >= 7:
            r['situação'] = 'BOA'
        elif r['média'] >= 5:
            r['situação'] = 'RAZOÁVEL'
        else:
            r['situação'] = 'RUIM'
    #Como colocamos um valor opcional (sit=False), colocamos ele pra analise, se foi ou não solicitado
    # Se o usuario chama o def notas(..., sit=True) ele chama a função sit junto a analise do dicionario;
    #Esse IF, apenas analisa a media e parametriza ela em BOA, RAZOAVEL ou RUIM;

    return r


#Programa Principal:
resp = notas(5.5,2.5,10,6.5, sit=True)
print(resp)
help(notas)


#Exercicio106: Faça um mini-sistema que utilize o Interactive Help do Python. 
#O usuário vai digitar o comando e o manual vai aparecer. Quando o usuário digitar a palavra ‘FIM’, o programa se encerrará. Importante: use cores.

cores = ('\033[m',          #0 - Sem cor
         '\033[0;30;41m',   #1 - Vermelho
         '\033[0;30;42m',   #2 - Verde
         '\033[0;30;43m',   #3 - Amarelo
         '\033[0;30;44m',   #4 - Azul
         '\033[0;30;45m',   #5 - Roxo
         '\033[7;30m',      #6 - Branco
         )


def ajuda(com):
    titulo('Acessando o manual do comando \'{comando}\'',4)

    print(cores[6], end='')
    help(com)
    print(cores[0], end='')


def titulo(msg,cor=0):
    tamanho_msg = len(msg)

    print(cores[cor], end='')

    print('~' * tamanho_msg)
    print(msg)
    print('~' * tamanho_msg)  

    print(cores[0], end='')  


#Programa principal:
comando = ''

while True:
    titulo('Sistema de ajuda PyHelp',2)

    comando = str(input('Função ou biblioteca > '))
    if comando.upper() == 'FIM':
        break
    else:
        ajuda(comando)
        

        
#Exercicio107: Crie um módulo chamado moeda.py que tenha as funções incorporadas aumentar(), diminuir(), dobro() e metade(). 
#Faça também um programa que importe esse módulo e use algumas dessas funções.
from modulos import ex107

resp = float(input('Digite um valor R$ '))

print(f'A metade do valor R${resp:.2F} é R$ {ex107.metade(resp)}')
print(f'O dobro do valor R${resp:.2F} é R$ {ex107.dobro(resp)}')
print(f'Aumentando em 10% o valor, temos R$ {ex107.aumentar(resp,10)}')
print(f'Diminuindo em 10% o valor, temos R$ {ex107.diminuir(resp,10)}')        


#Exercicio108: Adapte o código do desafio #107, criando uma função adicional chamada moeda() que consiga mostrar os números como um valor monetário formatado.
from modulos import ex108

resp = float(input('Digite um valor R$ '))

print(f'A metade do valor {ex108.moeda(resp)} é {ex108.moeda(ex108.metade(resp))}')
print(f'O dobro do valor {ex108.moeda(resp)} é {ex108.moeda(ex108.dobro(resp))}')
print(f'Aumentando em 10% o valor, temos {ex108.moeda(ex108.aumentar(resp,10))}')
print(f'Diminuindo em 10% o valor, temos {ex108.moeda(ex108.diminuir(resp,10))}')


#Exercicio109: Modifique as funções que form criadas no desafio 107 para que elas aceitem um parâmetro a mais, 
#informando se o valor retornado por elas vai ser ou não formatado pela função moeda(), desenvolvida no desafio 108.
from modulos import ex109

resp = float(input('Digite um valor R$ '))

print(f'A metade do valor {ex109.moeda(resp)} é {ex109.metade(resp,True)}')
print(f'O dobro do valor {ex109.moeda(resp)} é {ex109.dobro(resp, True)}')
print(f'Aumentando em 10% o valor, temos {ex109.aumentar(resp,10, True)}')
print(f'Diminuindo em 10% o valor, temos {ex109.diminuir(resp,10, True)}')


#Exercicio110: Adicione o módulo moeda.py criado nos desafios anteriores, 
#uma função chamada resumo(), que mostre na tela algumas informações geradas pelas funções que já temos no módulo criado até aqui.
from modulos import ex110

resp = float(input('Digite um valor R$ '))
ex110.resumo(resp,10,30)


#Exercicio111: Crie um pacote chamado utilidadesCeV que tenha dois módulos internos chamados moeda e dado. 
#Transfira todas as funções utilizadas nos desafios 107, 108 e 109 para o primeiro pacote e mantenha tudo funcionando.

""" Nesse exercicio apenas importamos de forma de pacote as informações, não alteramos nenhum codigo,
-No caso, montamos um pacote.python, e separamos o assunto pos pastas, dentro de cada pasta o codigo fica no __init__.py
é obrigatorio ter esse __init__ para cada pasta contendo as funções;
"""


#Exercicio112: Dentro do pacote utilidadesCeV que criamos no desafio 111, temos um módulo chamado dado. 
#Crie uma função chamada leiaDinheiro() que seja capaz de funcionar como a função imputa(), mas com uma validação de dados para aceitar apenas valores que seja monetários.
from modulos import ex112 #Nesse meu caso, não utilizei um pacote, fiz apenas o modulo ex112

resp = ex112.leiaDinheiro('Digite um valor R$ ')
ex112.resumo(resp,10,30)

#Exercicio113: Reescreva a função leiaInt() que fizemos no desafio 104, 
#incluindo agora a possibilidade da digitação de um número de tipo inválido. Aproveite e crie também uma função leiaFloat() com a mesma funcionalidade.
from modulos import ex113

n1 = ex113.leiaInt('Digite um valor: ')
n2 = ex113.leiaFloat('Digite um Real: ')
print(f'\033[0;32mVocê digitou o número inteiro {n1} e o número real {n2}.\033[m')

#Exercicio114: Crie um código em Python que teste se o site pudim está acessível pelo computador usado.
import urllib
import urllib.request
#Bibliotecas para importação de links

#Tratamento de erros
try: #Tente fazer isso:
    site = urllib.request.urlopen('https://www.pudim.com.br')

except urllib.error.URLError: #Se der erro no acesso a URL, retorne isso:
    print('\033[0;31mNão consegui acessar o site.\033[m') #Mensagem de erro!

else: #Se não der erro, retorne:
    print('\033[0;32mConsegui acessar o site Pudim com sucesso;.\033[m')
    #print(site.read) - Esse site.read, vai retornar todo o html da pagina
