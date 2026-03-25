#Levantamento de quantos alunos ficaram em cada situação.
aprovados = 0
recuperacao = 0
reprovados = 0 
ava1 = 0
ava2 = 0
ava3 = 0

#Criando um loop para escolher quantos alunos terão as médias avaliadas

rep = int(input('Quantos Alunos terão a média avaliada? '))

for i in range(rep):

#Receber Notas Dos ALunos

 n1 = float(input('Primeira Nota:  '))
 while n1 <0 or n1 >10:
    print('digite uma nota válida')
    n1 = float(input('Primeira Nota:  '))
 ava1 = ava1 + n1
 n2 = float(input('Segunda Nota:  '))
 while n2 <0 or n2 >10:
    print('digite uma nota válida')
    n2 = float(input('Segunda Nota:  '))
 ava2 = ava2 + n2
 n3 = float(input('Terceira Nota:  '))
 while n3 <0 or n3 >10:
    print('digite uma nota válida')
    n3 = float(input('Terceira Nota:  '))
 ava3 = ava3 + n3

#Declarar a variavel média e fazer o calculo da mesma
 media = (n1 + n2 + n3) / 3 

#Confirmar conforme a média se o aluno foi aprovado, reprovado ou está de recuperação
 if media >= 7:
    print(f'A média foi {media:.2f}, aluno aprovado.')
    aprovados = aprovados + 1

 elif 5<= media <7:
    print(f'A média foi {media:.2f}, aluno fará recuperação.')
    recuperacao = recuperacao + 1

 else:
    print(f'A média foi {media:.2f}, aluno reprovado.')
    reprovados = reprovados + 1




#Mostra quantos alunos ficaram em cada situação
print('Alunos aprovados: ', aprovados)
print('Alunos de recuperação: ', recuperacao)
print('Alunos reprovados: ', reprovados)
print('Média da avaliação 1:', format(ava1 / rep, '.2f'))
print('Média da avaliação 2:', format(ava2 / rep, '.2f'))
print('Média da avaliação 3:', format(ava3 / rep, '.2f'))

  
      

