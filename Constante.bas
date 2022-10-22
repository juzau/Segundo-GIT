Attribute VB_Name = "Módulo1"
Sub Constantes()
Const A1 As String = "A1"
Const A2 As String = "A2"

Dim nome As String
Dim numero As Integer

nome = InputBox("Digite o seu nome")
numero = InputBox("Digite um número")

Range(A1).Value = nome

If (numero Mod 2 = 0) Then
Range(A2).Value = "Este número é par"
Else
Range(A2).Value = "Este número é ímpar"
End If



End Sub



Sub mediaescolar()
Const media_aprovacao As Double = 7

'Para notas maiores ou iguais a 7: Aprovado
'Para notas menores ou iguais a 4: Reprovado
'Notas entre 4 e 7: Recuperação

Dim nota As Double
nota = InputBox("Digite a nota do aluno.")
If (nota > 10 Or nota < 0) Then
MsgBox "Nota inválida"

Else

    If (nota >= media_aprovação) Then
    MsgBox "Aprovado!"
    ElseIf (nota <= 4) Then
    MsgBox "Reprovado!"
    Else
    MsgBox "Recuperação"
    End If

End If

End Sub

