'INPUTbox pode ter até 7 parâmetros
'PROMPT: mensagem que aparece no centro da caixa de inserção
'TITLE: mensagem que aparece na barra de titulo
'DEFAULT: valor pré-definido da caixa de inserção (se o operador não digitar nada, vai ter a informação de parametro default inserido)
'xpos:  posição x de coordenada carteziana em que a caixa de inserção aparecerá
'ypos: posição y de coordenada carteziana em que a caixa de inserção aparecerá
'helpfile: arquivo de ajuda
'contexto: contexto do arquivo de ajuda

'exemplo de inputbosx:
'inputbox "Informe seu nome"

'introduzido variaveis
'nome = inputbox ("Informe seu nome")
'mgbox nome

nome = inputbox ("Informe seu nome", "Cadastro de pessoa", "seu nome", 0, 3000)
msgbox nome

