
Private Sub Workbook_Open()

Planilha1.Select

Call BloquearAbasComSenha

TelaCheia

    'Application.OnKey "{ESC}", ""
    'Application.OnKey "^{PGUP}", ""
    'Application.OnKey "^{PGDN}", ""
    
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    'C?digo a ser executado quando uma nova aba ? ativada
    Call Resolu
End Sub
