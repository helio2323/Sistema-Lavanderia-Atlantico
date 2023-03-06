Attribute VB_Name = "Functs"
    Private Declare PtrSafe Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Function converterNumeroEuropeu(ByVal numero As String) As String
    converterNumeroEuropeu = Replace(numero, ".", ",")
End Function


Function IsString(ByVal TestString As Variant) As Boolean
    If VarType(TestString) = vbString Then
        IsString = True
    Else
        IsString = False
    End If
End Function
Function converterNumero(ByVal numeroTexto As String) As Integer
    converterNumero = Val(numeroTexto)
End Function


Sub Atualizar()

'Application.ScreenUpdating = False

Call DesbloquearAbasComSenha
    ActiveWorkbook.RefreshAll
Call BloquearAbasComSenha

Range("e1").Select
Planilha6.Select

'Application.ScreenUpdating = True

End Sub

Sub Area_Cupom()

    Dim Area As Range
    Dim fim As Long
    Dim Resp As Variant
    
    
    Resp = MsgBox("Imprimir cupom não fiscal?", vbYesNo, "Atenção")
    
    If Resp = vbYes Then
    
        Sheets("Cupom").Select
        
        fim = Cells(1, 17).Value ' Ajuste para obter o valor numérico
        
        Set Area = Range("A1:E" & fim) ' Ajuste para definir diretamente o objeto Range
        
        ActiveSheet.PageSetup.PrintArea = "" ' redefine a área de impressao
        
        ActiveSheet.PageSetup.PrintArea = Area.Address ' Ajuste para atribuir o endereço da área
        
        On Error GoTo Erro22
        
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
        
    End If
        
Sheets("Pedido_Novo").Select

Exit Sub

Erro22: Toast "Atenção", "Sua impressora não foi instalada corretamente", 3

Sheets("Pedido_Novo").Select

End Sub

Sub Area_Etiqueta()

    Dim Area As Range
    Dim fim As Long
    Dim Resp As Variant
    
    
    Resp = MsgBox("Imprimir etiquetas?", vbYesNo, "Atenção")
    
    If Resp = vbYes Then
    
        Sheets("Etiqueta").Select
        
        fim = Cells(1, 10).Value ' Ajuste para obter o valor numérico
        
        Set Area = Range("A1:E" & fim) ' Ajuste para definir diretamente o objeto Range
        
        ActiveSheet.PageSetup.PrintArea = "" ' redefine a área de impressao
        
        ActiveSheet.PageSetup.PrintArea = Area.Address ' Ajuste para atribuir o endereço da área
        
        On Error GoTo Erro22
        ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
            IgnorePrintAreas:=False
        
    End If
        
Sheets("Pedido_Novo").Select

Exit Sub

Erro22: Toast "Atenção", "Sua impressora não foi instalada corretamente", 3

Sheets("Pedido_Novo").Select

        
End Sub



Function Dashboard()
    
Application.ScreenUpdating = False
    ativa = ActiveSheet.name
    
    Planilha4.Select
    
    VerificaUsuario ativa

Application.ScreenUpdating = True

End Function

Sub Resolu()


    Dim cyScreen As Long

    cxScreen = GetSystemMetrics(SM_CXSCREEN)
    cyScreen = GetSystemMetrics(SM_CYSCREEN)

    GetScreenResolution = cxScreen & "x" & cyScreen
    
    If cxScreen = 1920 Then
        ActiveWindow.Zoom = 110
    End If
    
    If cxScreen = 1366 Then
        ActiveWindow.Zoom = 80
    End If
    
    If cxScreen = 1024 Then
        ActiveWindow.Zoom = 70
    End If
    
    If ActiveSheet.name = "Pedido_Novo" Then
        ActiveWindow.Zoom = 80
    End If
    

End Sub

Function Produtividade()
Application.ScreenUpdating = False
    ativa = ActiveSheet.name
    
    Planilha20.Select
    
    VerificaUsuario ativa
Application.ScreenUpdating = True
End Function


Function VerificaUsuario(ativa)

Dim plan_ativa As String

plan_ativa = ativa

i = 7

Do

    If plan_ativacao.Cells(i, 7) = "Logado" Then
    If plan_ativacao.Cells(i, 6) = "padrao" Then
    
        If ActiveSheet.name = "Painel" Or ActiveSheet.name = "Produtividade" Then
        
                Toast "Atenção", "Você não tem permissão para acessar essa aba", 3
                
                Sheets(plan_ativa).Select
                Application.ScreenUpdating = True
                Exit Function
        
        End If
    
    End If
    End If
    
    i = i + 1
Loop Until plan_ativacao.Cells(i, 6) = ""


End Function




Function SpeedOn()

    With Application
        .Calculation = xlCalculationManual
        .ScreenUpdating = False
        .EnableEvents = False
        .DisplayAlerts = False
        .Cursor = xlWait
    End With
    
End Function
Function SpeedOff()

    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .EnableEvents = True
        .DisplayAlerts = True
        .Cursor = xlDefault
    End With
    
End Function

Function Select_Cx(Cx_Name)

ActiveSheet.Shapes.Range(Array(Cx_Name)).Select
Select_Cx = Selection.ShapeRange(1).TextFrame2.TextRange.Characters.text

End Function

Sub TelaCheia()

Application.DisplayFullScreen = True
Application.DisplayFormulaBar = False
ActiveWindow.DisplayFormulas = False
ActiveWindow.DisplayHeadings = False
ActiveWindow.DisplayWorkbookTabs = False



End Sub

Sub TelaFechada()

Application.DisplayFullScreen = False
Application.DisplayFormulaBar = True
'ActiveWindow.DisplayFormulas = True
ActiveWindow.DisplayHeadings = True
ActiveWindow.DisplayWorkbookTabs = True


End Sub
Function ToastGren(Alert_Ok As String)

Call MessageFormOpen(frmToastr.lblMsgIcon, Alert_Ok, 9095196)

End Function

Function ToastYelow(Alert_Atention As String)

Call MessageFormOpen(frmToastr.lblMsgIcon, Alert_Atention, 4113142)

End Function

Function Toast_Info(Alert_I As String)

Call MessageFormOpen(frmToastr.lblMsgIcon, Alert_I, 8880899)

End Function
Function TxtFormat_LostFocus()
    
    
    Cor01 = RGB(173, 184, 204)
    Cor02 = RGB(30, 30, 42)
    
    Dim ws As Worksheet
    Dim wa As Worksheet
    
    Dim Txt As MSForms.TextBox
    Dim Txt_Infos As Variant
    Dim X As Long
    Dim y As Long
    Dim z As Long
    
    Dim ShetName As String
    
    ShetName = ActiveSheet.name
    
    Set ws = ThisWorkbook.Worksheets(ShetName)
    Set wa = ThisWorkbook.Worksheets("TextBoxs")
    
    To_End = Last_Item(wa)
    
    X = 1
    y = 2
    z = 5
    
    For i = 2 To To_End
    
        'Carrega array da linha
        Txt_Infos = Array_Txt(wa, i)
        If Txt_Infos(z) = ShetName Then
            Set Txt = ws.OLEObjects(Txt_Infos(X)).Object
            
            With Txt

                If .Value = "" Then
                    .Value = Txt_Infos(y)
                    .ForeColor = Cor01
                End If
                If .Value <> Txt_Infos(y) Then
                    .ForeColor = Cor02
                End If
            
            End With
        End If
    Next i
    
TxtFormat_LostFocus = Txt.Value
    
End Function
Function TxtFormat_Focus()
    
    Cor01 = RGB(173, 184, 204)
    Cor02 = RGB(30, 30, 42)
    
    Dim ws As Worksheet
    Dim wa As Worksheet
    
    Dim Txt As MSForms.TextBox
    Dim Txt_Infos As Variant
    Dim X As Long
    Dim y As Long
    Dim z As Long
    
    Dim ShetName As String
    
    ShetName = ActiveSheet.name
    
    Set ws = ThisWorkbook.Worksheets(ShetName)
    Set wa = ThisWorkbook.Worksheets("TextBoxs")
    
    To_End = Last_Item(wa)
    
    X = 1
    y = 2
    z = 5
    
    For i = 2 To To_End
    
        'Carrega array da linha
        Txt_Infos = Array_Txt(wa, i)
        If Txt_Infos(z) = ShetName Then
            Set Txt = ws.OLEObjects(Txt_Infos(X)).Object
            
            With Txt
                If .Value = Txt_Infos(y) Then
                    .Value = ""
                    .ForeColor = Cor02
                End If
                If .Value <> Txt_Infos(y) Then
                    .ForeColor = Cor02
                End If
            
            End With
        End If
    Next i
    
TxtFormat_Focus = Txt.Value
    
End Function

Function Last_Item(Shet As Worksheet)

LastSale = Shet.Cells(Rows.Count, "A").End(xlUp).Row

Last_Item = LastSale

End Function

Function Array_Txt(Shet As Worksheet, Lin)

Dim arr(1 To 5)

For i = 1 To 5

    arr(i) = Shet.Cells(Lin, i)

Next i

Array_Txt = arr

'1 - Objeto
'2 - Placeholder
'3 - Tipo (String, Date, Bol
'4 - Obrigatorio (yes, no)


End Function
Function FormatEuro(Valor As Double) As String
    FormatEuro = Format(Valor, "€ #,##0.00")
End Function

Function RemoveLeadingZeros(text As String) As String
    Dim i As Long
    For i = 1 To Len(text)
        If Mid(text, i, 1) <> "0" Then
            RemoveLeadingZeros = Mid(text, i, Len(text) - i + 1)
            Exit Function
        End If
    Next i
End Function

Function ParOuImpar(numero As Integer) As String
    If numero Mod 2 = 0 Then
        ParOuImpar = "Par"
    Else
        ParOuImpar = "Ímpar"
    End If
End Function


