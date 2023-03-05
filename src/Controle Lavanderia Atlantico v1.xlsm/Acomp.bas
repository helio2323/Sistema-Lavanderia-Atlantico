Attribute VB_Name = "Acomp"
Sub Carrega_Acompanhamento(Pedido)
'SpeedOn
Dim y As Long

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set bd_pedidos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

wa.N__Pedido.Value = Pedido
On Error Resume Next
wa.Range("Lista_Itens").Delete
On Error GoTo 0

wa.Status_Resp.Value = ""
wa.Responsavel.Value = ""

y = 25
With bd_pedidos
    For i = 1 To bd_pedidos.Rows.Count

        If .Cells(i, 1) = Pedido Then
            
            wa.Status_Resp.Value = .Cells(i, 12)
            wa.Responsavel.Value = .Cells(i, 13)
            wa.Pagamento.Value = .Cells(i, 16)
            
            wa.Cells(y, 16) = .Cells(i, 7)
            wa.Cells(y, 17) = .Cells(i, 8)
            wa.Cells(y, 18) = .Cells(i, 9)
            wa.Cells(y, 19) = .Cells(i, 10)
            wa.Cells(y, 20) = .Cells(i, 11)
            y = y + 1
        End If
        
    Next i
End With

'SpeedOff

End Sub

Sub Carrega_Combo()

Dim wf As Worksheet

Set wf = ActiveWorkbook.Worksheets("Funcionários")

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set bd_pedidos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)



With wa
    
    .Status_Resp.Clear
    .Responsavel.Clear
    .Pagamento.Clear
    
    .Status_Resp.AddItem "Em Andamento"
    .Status_Resp.AddItem "Aguardando Retirada"
    .Status_Resp.AddItem "Entregue"
    
    For i = 1 To wf.Range("Funcionarios").Rows.Count
    
        .Responsavel.AddItem wf.Range("Funcionarios").Cells(i, 2)
    
    Next i
    
    .Pagamento.AddItem "Aguardando Pagamento"
    .Pagamento.AddItem "Pago"
    
    
End With

End Sub


Sub Atualiza_Status()

Dim y As Long

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set bd_pedidos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

Pedido = Acom.L_N_Pedido.Caption


y = 25
With bd_pedidos
    For i = 1 To bd_pedidos.Rows.Count
        Dado = .Cells(i, 1)
        If .Cells(i, 1) = CInt(Pedido) Then
            
            .Cells(i, 12) = Acom.Cb_Status.Value
            .Cells(i, 13) = Acom.Cb_Responsavel.Value
            .Cells(i, 16) = Acom.Cb_Pagamento.Value
            
            y = y + 1
            
        End If
        
    Next i
End With

Acom.Hide

DesbloquearAbasComSenha

ActiveSheet.PivotTables("TB_Acompanhamento").PivotCache.Refresh

BloquearAbasComSenha

Toast "Atenção!", "Dados atualizado com sucesso!", 1

End Sub

Sub Alterar_Acompanhamento()

    Carregar_Alteracao Acom.L_N_Pedido.Caption, True

End Sub

Sub preencher_tabela()

'SpeedOn

Dim i As Integer

For i = 58 To 5000 ' altere os valores de acordo com a sua tabela

    ' Preenche o número do pedido
    Cells(i, 1) = i - 1
    
    ' Preenche a data
    Cells(i, 2) = Date + Int(Rnd() * 100)
    
    ' Preenche o levantamento
    Cells(i, 3) = Date + Int(Rnd() * 10)
    
    ' Preenche o documento
    Cells(i, 4) = Int(Rnd() * 100000000)
    
    ' Preenche o nome
    Cells(i, 5) = "Nome " & i
    
    ' Preenche a morada
    Cells(i, 6) = "Morada " & i
    
    ' Preenche o tipo de serviço
    Cells(i, 7) = "Serviço " & i
    
    ' Preenche a unidade
    Cells(i, 8) = "UN"
    
    ' Preenche o valor
    Cells(i, 9) = Int(Rnd() * 100)
    
    ' Preenche a quantidade
    Cells(i, 10) = Int(Rnd() * 10)
    
    ' Preenche o total
    Cells(i, 11) = Cells(i, 9) * Cells(i, 10)
    
    ' Preenche o status interno
    Cells(i, 12) = "Entregue"
    
    ' Preenche o responsável
    Cells(i, 13) = "Responsável " & i
    
    ' Preenche o prazo em dias
    Cells(i, 14) = Int(Rnd() * 10) - 2


Next i

'SpeedOff

Toast


End Sub


