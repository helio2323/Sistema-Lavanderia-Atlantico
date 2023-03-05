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

Private Sub Worksheet_Activate()
    Call CRUD_Read
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)

End Sub