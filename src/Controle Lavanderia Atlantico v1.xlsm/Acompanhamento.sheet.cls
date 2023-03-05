Private Sub Pedidos_List_Click()

End Sub

Private Sub Pedidos_List_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Linha As Long

    Linha = Me.Pedidos_List.ListIndex 'captura o índice da linha selecionada
    
    Carregar_Alteracao (Linha)
    
End Sub

Private Sub Pedidos_Paginador_Change()
    
End Sub

Private Sub Pedidos_Paginador_LostFocus()
    CRUD_Read
End Sub

Private Sub Pesquisa_GotFocus()
    TxtFormat_Focus
End Sub
Private Sub Pesquisa_LostFocus()
    TxtFormat_LostFocus
End Sub
Private Sub Txt1_GotFocus()
    TxtFormat_Focus
End Sub

Private Sub Txt1_LostFocus()
    TxtFormat_LostFocus
End Sub



Private Sub Pagamento_Change()

End Sub

Private Sub Pagamento_LostFocus()
    Atualiza_Status
End Sub

Private Sub Responsavel_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Status_Resp.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Responsavel_LostFocus()
    Atualiza_Status
End Sub


Private Sub Status_Resp_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Responsavel.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Status_Resp_LostFocus()
    Atualiza_Status
End Sub

Private Sub Worksheet_Activate()

End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Target.Column = 6 And Target.Row > 1 Then
        'A coluna F foi selecionada
        Call Carrega_Dados_Acompanhamento
        'Toast 'Substitua "SuaMacro" pelo nome da sua macro
    End If
End Sub
