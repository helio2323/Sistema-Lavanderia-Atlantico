



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

Private Sub Data_Change()
    



End Sub

Private Sub Data_GotFocus()

End Sub

Private Sub Data_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Documento.Activate
        KeyCode = 0
    End If

End Sub

Private Sub Data_LostFocus()

End Sub

Private Sub Documento_Change()

End Sub

Private Sub Documento_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Nome.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Levantamento_Change()

End Sub

Private Sub Levantamento_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Data.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Lista_Pedido_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Linha As Long

    Linha = Me.Lista_Pedido.ListIndex 'captura o índice da linha selecionada
    
    Delete_Pedido (Linha + 1)
    
End Sub


Private Sub Modelo_Change()

End Sub

Private Sub Modelo_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Tipo_Servico.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Morada_Change()

End Sub

Private Sub Morada_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Telefone.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Nome_Change()

End Sub

Private Sub Nome_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Morada.Activate
        KeyCode = 0
    End If
End Sub



Private Sub OptionButton1_Click()

End Sub

Private Sub Pagamento_Change()
Planilha14.Cells(1, 22) = Planilha9.Pagamento.Value
End Sub

Private Sub Pagamento_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Modelo.Activate
        KeyCode = 0
    End If
End Sub



Private Sub Pagamento_LostFocus()
    
End Sub

Private Sub Quantidade_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Levantamento.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Telefone_Change()

End Sub

Private Sub Telefone_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Pagamento.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Tipo_Servico_Change()

End Sub

Private Sub Tipo_Servico_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        If wa.Tipo_Servico.Value = "Outros Serviços" Then
            wa.Valor.Activate
            KeyCode = 0
        Else
            wa.Quantidade.Activate
            KeyCode = 0
        End If
    End If
End Sub

Private Sub Tipo_Servico_LostFocus()
    Carrega_Valores
End Sub

Private Sub Unidade_Change()
    
End Sub


Private Sub valor_Change()

End Sub



Private Sub valor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    Objetos = GetListboxVariables()
    
    Set wa = Objetos(6)
    
    If KeyCode = vbKeyTab Then
        wa.Quantidade.Activate
        KeyCode = 0
    End If
End Sub

Private Sub Worksheet_Activate()
    Call Carrega_ComboBox
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub