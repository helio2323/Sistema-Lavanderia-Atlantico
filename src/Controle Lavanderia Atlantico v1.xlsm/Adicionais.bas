Attribute VB_Name = "Adicionais"
Sub Salva_Plan()
    
    ActiveWorkbook.Save
    
    Toast "Documento Salvo", "Sua planilha foi salva", 1

End Sub
Sub Cli()

   Clientes.Show

End Sub

Sub calendario()

    Data = GetCalend�rio
    
    Planilha9.Data.Value = Data

End Sub

Sub calendario_2()

    Data = GetCalend�rio
    
    Planilha9.Levantamento.Value = Data
    
    Planilha14.Cells(1, 20) = Data

End Sub

Sub Servs()

   servicos.Show

End Sub

Sub Carrega_Dados_Acompanhamento()
    
    Dim Resp As Variant
    
    If ActiveCell.Value2 = "" Then
        Exit Sub
    End If
    
    Valor = ActiveCell.Value2
    
    If IsString(Valor) Then
        Exit Sub
    End If
    
    
    Resp = MsgBox("Voc� quer carregar o pedido " & ActiveCell.Value2 & "?", vbYesNo, "Aten��o")
    
    If Resp = vbNo Then
        Exit Sub
    End If
    
    
    Acom.Show

End Sub
