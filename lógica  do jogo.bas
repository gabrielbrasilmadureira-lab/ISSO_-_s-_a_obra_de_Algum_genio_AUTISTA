Attribute VB_Name = "Módulo1"
Dim TimeA As Integer
Dim TimeB As Integer
Dim Segundos As Integer

Sub IniciarPartida()
    TimeA = 0
    TimeB = 0
    Segundos = 0
    
    ' Inicianiliza o gerador aleatório
    Randomize
    
    
    'Inicia o crônometro
    Call RodarCronometro
End Sub

 Sub RodarCronometro()
    Dim i As Integer
    
    For i = 1 To 45
    Segundos = i
    
    'Simula chance de gol para cada time
    If Rnd < 0.1 Then TimeA = TimeA + 1
    If Rnd < 0.1 Then TimeB = TimeB + 1
    
    
    'Atualiza na planilha
    Cells(1, 1).Value = "Tempo: " & Segundos & "s"
    Cells(2, 1).Value = "Time A: " & TimeA
    Cells(3, 1).Value = "Time B: " & TimeB
    
    ' Espera 1 segundo
    Application.Wait (Now + TimeValue("0:00:01"))
    DoEvents


Next i


MsgBox "Fim do primeiro Tempo! Placar: " & TimeA & " x " & TimeB
End Sub
