' VBScript para contar a história interativa de Joãozinho e Margarida

Option Explicit

' Variáveis globais
Dim escolha, continuar
escolha = ""
continuar = True

' Função principal que inicia a história
Sub IniciarHistoria()
    MsgBox "Era uma vez, em uma pequena vila cercada por uma densa floresta, viviam Joãozinho e sua irmã Margarida."
    MsgBox "Um dia, seus pais, afligidos pela fome, decidiram levá-los para a floresta e abandoná-los lá."
    MsgBox "As crianças, temerosas e confusas, viram-se perdidas no meio da floresta."
    EscolherCaminho()
End Sub

' Função para permitir que o jogador faça uma escolha
Sub EscolherCaminho()
    continuar = True
    Do While continuar
        escolha = InputBox("O que Joãozinho e Margarida devem fazer? (Digite A, B ou C)" & vbCrLf & _
                           "A) Espalhar Migalhas de Pão" & vbCrLf & _
                           "B) Guardar o Pão para Comer Depois" & vbCrLf & _
                           "C) Tentar Convencer os Pais a Não Abandoná-los")
                           
        If UCase(escolha) = "A" Then
            EspalharMigalhasDePao()
        ElseIf UCase(escolha) = "B" Then
            GuardarPaoParaComerDepois()
        ElseIf UCase(escolha) = "C" Then
            TentarConvencerOsPais()
        Else
            MsgBox "Escolha inválida. Por favor, escolha A, B ou C."
        End If
    Loop
End Sub

' Opção A: Espalhar Migalhas de Pão
Sub EspalharMigalhasDePao()
    MsgBox "Joãozinho espalhou as migalhas de pão pelo caminho, esperando que elas os guiassem de volta para casa."
    MsgBox "No entanto, as aves da floresta comeram as migalhas, deixando-os completamente perdidos."
    continuar = False
End Sub

' Opção B: Guardar o Pão para Comer Depois
Sub GuardarPaoParaComerDepois()
    MsgBox "Joãozinho decidiu guardar o pão para ele e Margarida comerem mais tarde."
    MsgBox "Sem um caminho marcado, eles se perderam na floresta e enfrentaram uma longa jornada."
    continuar = False
End Sub

' Opção C: Tentar Convencer os Pais a Não Abandoná-los
Sub TentarConvencerOsPais()
    MsgBox "Joãozinho tentou convencer seus pais a não os deixarem na floresta, mas seus apelos foram em vão."
    MsgBox "Desesperados e com o coração partido, Joãozinho e Margarida foram abandonados na floresta escura."
    continuar = False
End Sub

' Início do script
Call IniciarHistoria()
