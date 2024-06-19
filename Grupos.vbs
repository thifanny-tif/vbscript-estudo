'MSGBOX é composto por até 5 parâmetros
'PROMPT: mensagem que aparece no centro da caixa de texto
'BUTTON: composto de 4 grupos de números
'	Grupo 1: Configuração de texto
'	Grupo 2: Ícones de caixa de texto
'	Grupo 3: Botão padrão
'	Grupo 4: Modal
'TITLE: mensagem aparece na barra de titulo da caixa de texto
'HELPTITLE: endereço do arquivo na barra de caixa de texto
'CONTEXT: contexto do arquivo de ajuda

'GRUPO 1

'Cenário 1
'Condições:
'-Ser informado o valor 0
'Resultado:
'- A caixa de texto exibe o botão "OK"

'Cenário 2
'Condições:
'-Ser informado o valor 1
'Resultado:
'- A caixa de texto exibe o botão "OK", "Cancelar"

'Cenário 3
'Condições:
'-Ser informado o valor 2
'Resultado:
'- A caixa de texto exibe o botão "Abortar", "Repetir" e "Cancelar"

'Cenário 4
'Condições:
'-Ser informado o valor 3
'Resultado:
'- A caixa de texto exibe o botão "Sim", "Não" e "Cancelar"

'Cenário 5
'Condições:
'-Ser informado o valor 4
'Resultado:
'- A caixa de texto exibe o botão "Sim" e "Não"

'Cenário 6
'Condições:
'-Ser informado o valor 5
'Resultado:
'- A caixa de texto exibe o botão "Repetir" e "Cancelar"


'GRUPO 2

'Cenário 1
'Condições:
'-Ser informado o valor 0
'Resultado:
'- A caixa de texto não exibirá icone

'Cenário 2
'Condições:
'-Ser informado o valor 16
'Resultado:
'- A caixa de texto exibirá ícone de erro

'Cenário 3
'Condições:
'-Ser informado o valor 32
'Resultado:
'- A caixa de texto exibirá ícone de questão

'Cenário 4
'Condições:
'-Ser informado o valor 48
'Resultado:
'- A caixa de texto exibirá ícone de exclamação

'Cenário 5
'Condições:
'-Ser informado o valor 64
'Resultado:
'- A caixa de texto exibirá ícone informativo


'GRUPO 3
'Cenário 1
'Condições:
'-Ser informado o valor 0
'Resultado:
'- o primeiro botão ficará focadd

'Cenário 2
'Condições:
'-Ser informado o valor 256
'Resultado:
'- O segundo botão ficara focado

'Cenário 3
'Condições:
'-Ser informado o valor 512
'Resultado:
'- '- O terceiro botão ficara focado

'Cenário 4
'Condições:
'-Ser informado o valor 768
'Resultado:
'- O quarto botão ficará focado

'GRUPO 4
'Cenário 1
'Condições:
'-Ser informado o valor 0
'Resultado:
'-'- MODAL desativado

'Cenário 2
'Condições:
'-Ser informado o valor 4096
'Resultado:
'- MODAL ativado

msgbox "olá mundo", 2+16+256+4096, "ATENÇÃO"

