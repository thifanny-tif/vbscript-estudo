Option Explicit

Dim operacao, A, B, resultado, continuar

Do
    ' Solicita ao usuário para escolher a operação
    operacao = InputBox("Escolha a operação: " & vbCrLf & "1 - Adição" & vbCrLf & "2 - Subtração" & vbCrLf & "3 - Multiplicação" & vbCrLf & "4 - Divisão" & vbCrLf & "5 - Raiz Quadrada" & vbCrLf & "6 - Sair")

    If operacao = "6" Then
        Exit Do
    End If

    Select Case operacao
        Case "1"
            ' Adição
            A = InputBox("Informe o primeiro número") + 0
            B = InputBox("Informe o segundo número") + 0
            resultado = A + B
            MsgBox "O resultado da adição é: " & resultado
        Case "2"
            ' Subtração
            A = InputBox("Informe o primeiro número") + 0
            B = InputBox("Informe o segundo número") + 0
            resultado = A - B
            MsgBox "O resultado da subtração é: " & resultado
        Case "3"
            ' Multiplicação
            A = InputBox("Informe o primeiro número") + 0
            B = InputBox("Informe o segundo número") + 0
            resultado = A * B
            MsgBox "O resultado da multiplicação é: " & resultado
        Case "4"
            ' Divisão
            A = InputBox("Informe o primeiro número") + 0
            B = InputBox("Informe o segundo número") + 0
            If B = 0 Then
                MsgBox "Erro: Divisão por zero não é permitida."
            Else
                resultado = A / B
                MsgBox "O resultado da divisão é: " & resultado
            End If
        Case "5"
            ' Raiz Quadrada
            A = InputBox("Informe o número para calcular a raiz quadrada") + 0
            If A < 0 Then
                MsgBox "Erro: Não é possível calcular a raiz quadrada de um número negativo."
            Else
                resultado = A ^ 0.5
                MsgBox "O resultado da raiz quadrada é: " & resultado
            End If
        Case Else
            MsgBox "Operação inválida. Por favor, escolha uma operação entre 1 e 5."
    End Select

    continuar = MsgBox("Deseja realizar outra operação?", vbYesNo + vbQuestion, "Calculadora")
    If continuar = vbNo Then
        Exit Do
    End If

Loop

MsgBox "Obrigado por usar a calculadora. Até mais!"
