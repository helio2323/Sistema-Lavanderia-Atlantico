Attribute VB_Name = "Seguranca"
Sub BloquearAbasComSenha()

    Dim Senha As String
    Dim aba As Worksheet
    Dim AbasBloqueadas(1 To 7)
    
    'Defina as abas que deseja bloquear
   
    
    For i = 1 To 7
    
        AbasBloqueadas(i) = Planilha2.Cells(i, 1)
    
    Next i
    
    
    
    'Defina a senha que será usada para desbloquear as abas
    Senha = "Helio@232425"
    
    'Bloqueie as abas selecionadas com a senha
    For Each aba In ActiveWorkbook.Worksheets
        If IsInArray(aba.name, AbasBloqueadas) Then
            aba.Protect Password:=Senha, DrawingObjects:=True, Contents:=True, Scenarios:=True _
        , AllowFiltering:=True, AllowUsingPivotTables:=True
        End If
    Next aba
    
End Sub

Function IsInArray(stringToBeFound As String, arr As Variant) As Boolean
    IsInArray = (UBound(Filter(arr, stringToBeFound)) > -1)
End Function

Sub DesbloquearAbasComSenha()

    Dim Senha As String
    Dim aba As Worksheet
    Dim AbasBloqueadas(1 To 13)
    
    'Defina as abas que deseja desbloquear
   
    
    For i = 1 To 13
    
        AbasBloqueadas(i) = Planilha2.Cells(i, 1)
    
    Next i
    
    
    
    'Defina a senha que será usada para desbloquear as abas
    Senha = "Helio@232425"
    
    'Desbloqueie as abas selecionadas com a senha
    For Each aba In ActiveWorkbook.Worksheets
        If IsInArray(aba.name, AbasBloqueadas) Then
            aba.Unprotect Password:=Senha
        End If
    Next aba
    
End Sub

Sub Salvar()

ActiveWorkbook.Save

Toast "Atenção", "Planilha salva com sucesso!", 1

End Sub





