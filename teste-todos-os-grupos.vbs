' Grupo 1: Botões exibidos na caixa de texto
' Grupo 2: Ícones exibidos na caixa de texto
' Grupo 3: Foco do botão
' Grupo 4: Modal ativado/desativado

Dim valorGrupo1, valorGrupo2, valorGrupo3, valorGrupo4
Dim combinacao

' Loop através de todas as combinações possíveis
For valorGrupo1 = 0 To 5
    For valorGrupo2 = 0 To 4
        For valorGrupo3 = 0 To 3
            For valorGrupo4 = 0 To 1
                ' Calcula o valor total para a combinação atual
                combinacao = valorGrupo1 + valorGrupo2 * 16 + valorGrupo3 * 256 + valorGrupo4 * 4096
                
                ' Mensagem a ser exibida
                Dim mensagem
                mensagem = "Combinação: Grupo1-" & valorGrupo1 & ", Grupo2-" & valorGrupo2 & ", Grupo3-" & valorGrupo3 & ", Grupo4-" & valorGrupo4
                
                ' Exibe a MsgBox com os parâmetros calculados
                MsgBox mensagem, combinacao, "Título da MsgBox"
            Next
        Next
    Next
Next
