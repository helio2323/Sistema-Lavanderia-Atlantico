Attribute VB_Name = "Pedidos"
Sub OcultarAbas()

Dim ws As Worksheet

For Each ws In ThisWorkbook.Sheets
    If ws.name <> "Login" And ws.name <> "Guia_Funcoes" Then
        ws.Visible = xlSheetVeryHidden
    End If
    
    Planilha1.Select
Next ws

End Sub

Sub teste()
Toast "HR Solutions", "Você acabou de salvar um item, parabens", l
ToastGren "Olá"
'ToastYelow "OI"

End Sub

Sub CRUD_Read()

Dim NPages As Long
Dim Txt_Pes_result As Variant

variables = GetListboxVariables()



Set ws = variables(0)
Set rg = variables(1) 'Range da tabela
Set Paginador = variables(2)
Set ListB = variables(3)
Set Txt_Pesquisa = variables(4)
ColPesquisa = variables(5)

NPages = 30
    
Pagi_Backup = Paginador.Value
    
Pages = Application.WorksheetFunction.RoundUp(rg.Rows.Count / NPages, 0)

Paginador.Clear

For i = 1 To Pages

    With Paginador
    
        .AddItem i
    
    End With

Next i

Paginador.Value = Pagi_Backup

If Paginador.Value = "" Then

    Paginador.Value = 1
    
End If

    Dim linf    As Integer
    Dim Cor01   As Variant
    Dim Cor02   As Variant
    Dim Cor03   As Variant

    Cor01 = RGB(35, 207, 222) 'RGB(253, 6, 100)
    Cor02 = RGB(43, 46, 51) 'RGB(231, 232, 237)
    
    ListB.ForeColor = Cor01

    ListB.BackColor = Cor02
    
    ListB.Clear

    ListB.ColumnHeads = False
    
    Star = 1 + (NPages * Paginador.Value) - NPages
    X = 1
    
    Total_De_Linhas = rg.Rows.Count
    
    ListB.ColumnWidths = "70;70;70;70;100;120;100;70;70;70;70;70"
    ListB.ColumnCount = rg.Columns.Count
    
    For linf = Star To NPages * Paginador.Value
    rg_value = rg.Cells(linf, 1).Value
    If rg_value <> Empty Then
        With ListB
            'parei aqui  ----------------------------------------------
            For i = 1 To 10

                If i = 1 Then
                    rg_value = Format(rg.Cells(linf, i).Value, "00000")
                    .AddItem rg_value
                Else
                    rg_value = rg.Cells(linf, i).Value
                    '.List(X - 1, i - 1) = rg_value
                    .List(.ListCount - 1, i - 1) = rg.Cells(linf, i).Value
                    If i = 9 Then
                         .List(.ListCount - 1, i - 1) = FormatEuro(rg.Cells(linf, i).Value)
                    End If
                End If
                    '.AddItem rg_value
            Next i
            
        End With
    X = X + 1
    End If
    Next
      
    'Me.Contalbl.Caption = Planilha6.ListBox1.ListCount & " clientes"

End Sub
Sub dd()

Dim ws As Worksheet
Dim Txt As MSForms.TextBox

Set ws = ThisWorkbook.Worksheets("Pedido_Cad")
Set Txt = ws.OLEObjects("Textbox1").Object

With Txt
    .Value = "CCC"
End With

End Sub



Function GetListboxVariables(Optional Linha As Variant = 0) As Variant
    Dim ws As Worksheet
    Dim rg As Range
    Dim Pages As Long

    Dim wa As Worksheet
    Dim ListT As Variant
    
    Dim ListBox_Configs As Worksheet
    Dim Aba_Ativa As String
    Dim Tabela As Range
    Dim Paginador As MSForms.ComboBox
    Dim PaginN As Variant
    Dim ListB As MSForms.ListBox
    Dim Last_Date As Long
    Dim Txt_Pesquisa As MSForms.TextBox
    Dim ColPesquisa As String
    Dim Txt_Name As String
    
    Set ListBox_Configs = ThisWorkbook.Worksheets("ListBox")
    
     
    Aba_Ativa = ThisWorkbook.ActiveSheet.name
    Set wa = ThisWorkbook.Worksheets(Aba_Ativa)
    
    
    For i = 2 To Last_Item(ListBox_Configs)
    
        If Aba_Ativa = ListBox_Configs.Cells(i, 1) Then
            If Linha > 0 Then
                i = Linha
            End If
            'Busca o nome do elemento na aba e define em uma variável
            ListT = ListBox_Configs.Cells(i, 4)
            PaginN = ListBox_Configs.Cells(i, 5)
            Txt_Name = ListBox_Configs.Cells(i, 6)
            ColPesquisa = ListBox_Configs.Cells(i, 7)
            
            'Define a variavel como um objeto
            Set ws = ThisWorkbook.Worksheets(ListBox_Configs.Cells(i, 2).Value)
            Set rg = ws.Range(ListBox_Configs.Cells(i, 3).Value)
            On Error Resume Next
            Set Paginador = wa.OLEObjects(PaginN).Object
            Set ListB = wa.OLEObjects(ListT).Object
            Set Txt_Pesquisa = wa.OLEObjects(Txt_Name).Object
            
            'Define a aba que está o banco de dados e a tabela
            GetListboxVariables = Array(ws, rg, Paginador, ListB, Txt_Pesquisa, ColPesquisa, wa)
            Exit Function
        End If
    
    Next i
    
    'Se não encontrar a configuração correta na tabela, retorna um erro
    Toast vbObjectError + 1000, "Não foi possível encontrar as variáveis do Listbox.", 2
   ' Err.Raise vbObjectError + 1000, "GetListboxVariables", "Não foi possível encontrar as variáveis do Listbox."
End Function


Sub Carrega_ComboBox()

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

backup_pagamento = wa.Pagamento
backup_modelo = wa.Modelo

Servicos_ComboBox.Clear
wa.Pagamento.Clear
wa.Modelo.Clear

For i = 1 To Tabela_Servicos.Rows.Count
    With Servicos_ComboBox
        .AddItem Tabela_Servicos(i, 1)
    End With
    Next i

    wa.Pagamento.AddItem "Aguardando Pagamento"
    wa.Pagamento.AddItem "Pago"
    
    wa.Modelo.AddItem "Entrega"
    wa.Modelo.AddItem "Levantamento"

wa.Unidade.Value = ""
wa.Valor.Value = ""
wa.Quantidade.Value = ""

wa.Pagamento = backup_pagamento
wa.Modelo = backup_modelo

End Sub

Sub Carrega_Valores()

Objetos = GetListboxVariables()

Set wa = Objetos(6)
Set Servicos_ComboBox = Objetos(2)
Set Tabela_Servicos = Objetos(1)

For i = 1 To Tabela_Servicos.Rows.Count
    
    With Servicos_ComboBox
        If .Value = Tabela_Servicos.Cells(i, 1) Then
            wa.Unidade = Tabela_Servicos.Cells(i, 2)
            wa.Valor = Tabela_Servicos.Cells(i, 3)

            If .Value = "Outros Serviços" Then
                wa.Unidade.Enabled = False
                wa.Valor.Enabled = True
            Else
                wa.Unidade.Enabled = False
                wa.Valor.Enabled = False
            End If
            Exit Sub
        End If
    End With
Next i

With Servicos_ComboBox
    wa.Unidade = "" 'verificar o porque da erro
    wa.Valor = ""

End With



End Sub

Sub Salva_Resumo_Pedido()

Dim Tb_Resumo_ As Range
Dim ws As Worksheet
Dim i As Variant

ActiveSheet.Shapes("Picture 10288").Visible = True

Objetos = GetListboxVariables()

Set wa = Objetos(6)
Set Resumo_Servicos = Objetos(3) 'Lista de serviços

Set ws = ThisWorkbook.Worksheets("Resumo_Pedido")
Set Tb_Resumo_ = ws.Range("TB_Resumo")

Add = Tb_Resumo_.Cells(1, 1)

If Add = "" Then
    Add = 0
Else
    Add = 1
End If

i = Tb_Resumo_.Rows.Count

If Tb_Resumo_.Cells(i, 1) <> "" Then
    i = i + 1
    On Error GoTo Erro22:
    Valor_Item = converterNumeroEuropeu(wa.Valor.Value)
    Tb_Resumo_.Cells(i, 5).Value = wa.Quantidade.Value * Valor_Item
    Tb_Resumo_.Cells(i, 1).Value = wa.Tipo_Servico.Value
    Tb_Resumo_.Cells(i, 2).Value = wa.Unidade.Value
    Tb_Resumo_.Cells(i, 3).Value = wa.Valor.Value
    Tb_Resumo_.Cells(i, 4).Value = wa.Quantidade.Value
    
    
    On Error GoTo 0
Else
    Tb_Resumo_.Cells(i, 1).Value = wa.Tipo_Servico.Value
    Tb_Resumo_.Cells(i, 2).Value = wa.Unidade.Value
    Tb_Resumo_.Cells(i, 3).Value = wa.Valor.Value
    Tb_Resumo_.Cells(i, 4).Value = wa.Quantidade.Value
    Quantidade_Item = Int(wa.Quantidade.Value)
    Valor_Item = converterNumeroEuropeu(wa.Valor.Value)
    Tb_Resumo_.Cells(i, 5).Value = wa.Quantidade.Value * Valor_Item
End If

Resumo_Servicos.Clear
Resumo_Servicos.ColumnWidths = "100;70;70;55;50;50"
Resumo_Servicos.ColumnCount = Tb_Resumo_.Columns.Count + 1


X = 0

For i = 1 To Tb_Resumo_.Rows.Count + Add
    With Resumo_Servicos
        .AddItem Tb_Resumo_.Cells(i, 1)
        .List(X, 2) = Tb_Resumo_.Cells(i, 2)
        .List(X, 3) = FormatEuro(Tb_Resumo_.Cells(i, 3))
        .List(X, 4) = Tb_Resumo_.Cells(i, 4)
        .List(X, 5) = FormatEuro(Tb_Resumo_.Cells(i, 5))
    End With
    X = X + 1
Next i

wa.Unidade.Value = ""
wa.Valor.Value = ""
wa.Quantidade.Value = ""
wa.Tipo_Servico.Value = ""

Call Cumpom_Fiscal

'ToastGren "Novo item adicionado - " & wa.Tipo_Servico.Value
Exit Sub

Erro22: MsgBox "Você digitou alguma letra ou não inseriu valor do produto e quantidade..", vbCritical
End Sub

Sub Delete_Pedido(Linha)

Dim Tb_Resumo_ As Range
Dim ws As Worksheet
Dim i As Variant
Dim Resp As Variant

Resp = MsgBox("Você selecionou a linha, gostária de excluir? " & Linha, vbYesNo, "Atenção")

If Resp = vbNo Then
    Exit Sub
End If

Objetos = GetListboxVariables()

Set wa = Objetos(6)
Set Resumo_Servicos = Objetos(3) 'Lista de serviços

Set ws = ThisWorkbook.Worksheets("Resumo_Pedido")
Set Tb_Resumo_ = ws.Range("TB_Resumo")

Tb_Resumo_.Rows(Linha).Delete

'Toast_Info "Item excluído - " & Tb_Resumo_.Cells(Linha, 1)

Resumo_Servicos.Clear
Resumo_Servicos.ColumnWidths = "100;50;50;50;50;50"
Resumo_Servicos.ColumnCount = Tb_Resumo_.Columns.Count + 1


X = 0

For i = 1 To Tb_Resumo_.Rows.Count
    With Resumo_Servicos
        .AddItem Tb_Resumo_.Cells(i, 1)
        .List(X, 2) = Tb_Resumo_.Cells(i, 2)
        .List(X, 3) = FormatEuro(Tb_Resumo_.Cells(i, 3))
        .List(X, 4) = Tb_Resumo_.Cells(i, 4)
        .List(X, 5) = FormatEuro(Tb_Resumo_.Cells(i, 5))
    End With
    X = X + 1
Next i

Call Cumpom_Fiscal

End Sub

Sub Salva_Pedido(Optional tipo As Boolean = False)

Application.ScreenUpdating = False

Dim TT_Itens As Variant
Dim Last_L As Variant
Dim ws As Worksheet
Dim Arr_Dados(1 To 8)
Dim Resp As Variant
Dim L_Ver As Variant
Dim wl As Worksheet

Objetos = GetListboxVariables(4)

Set wl = ThisWorkbook.Worksheets("TextBoxs")

L_Ver = Last_Item(wl)

Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set wa = Objetos(6)
Set ws = ThisWorkbook.Worksheets("Resumo_Pedido")

ped = Planilha9.N_Pedido.Value

'Faz uma verificação nas textbox, caso esteja vazio para o código
    For i = 2 To L_Ver

    aba = wl.Cells(i, 5)
    Obrigatorio = wl.Cells(i, 4)
    
    If aba = wa.name And Obrigatorio = "yes" Then
    
        Txt = wl.Cells(i, 1)
        Padrao = wl.Cells(i, 2)
        Valor = wa.OLEObjects(Txt).Object
        
        If Valor = "" Or Valor = Padrao Then
            
            MsgBox "Você precisa preencher: " & Txt
            Toast "Atenção", "Você precisa preencher: " & Txt, 3
            Exit Sub

            
        End If
        
    
    End If

Next i


If tipo = False Then
    Resp = MsgBox("Clique em 'Sim' para confirmar a inclusão do pedido: " & wa.N_Pedido, vbYesNo, "Atenção")
End If

If Resp = vbNo Then
    Exit Sub
End If


TT_Itens = Last_Item(ws)
Last_L = Tabela_Servicos.Rows.Count

If Tabela_Servicos.Cells(Last_L, 7) <> "" Then
    Last_L = Last_L + 1
End If

Arr_Dados(1) = wa.N_Pedido.Value
Arr_Dados(2) = wa.Data.Value
Arr_Dados(4) = wa.Documento.Value
Arr_Dados(5) = wa.Nome.Value
Arr_Dados(6) = wa.Morada.Value
Arr_Dados(3) = wa.Levantamento.Value
Arr_Dados(7) = wa.Pagamento.Value
Arr_Dados(8) = wa.Modelo.Value

Planilha14.Cells(1, 22).Value = Arr_Dados(7)

c = 1

For i = 2 To TT_Itens
    
    For c = 1 To 6
        Tabela_Servicos.Cells(Last_L, c + 6) = ws.Cells(i, c)
        Tabela_Servicos.Cells(Last_L, c) = Arr_Dados(c)
    Next c
    
    Tabela_Servicos.Cells(Last_L, 16) = Arr_Dados(7)
    Tabela_Servicos.Cells(Last_L, 17) = Arr_Dados(8)
    Tabela_Servicos.Cells(Last_L, 22) = Planilha9.Cells(43, 6)
    Tabela_Servicos.Cells(Last_L, 23) = Planilha9.Cells(43, 7)
    Tabela_Servicos.Cells(Last_L, 24) = Planilha9.Cells(43, 8)

    Last_L = Last_L + 1
    
Next i

'Limpa a tabela de resumos
On Error GoTo SemItens
ws.Range("TB_Resumo").Delete
On Error GoTo 0
If tipo = False Then
    Toast "Atenção!", "Você fez a inclusão do pedido: " & wa.N_Pedido, 1
Else
    Toast "Atenção!", "Você fez a alteração do pedido: " & wa.N_Pedido, 1
End If

Call Atualiza_Clentes

wa.Lista_Pedido.Clear

wa.N_Pedido.Value = ""
wa.Data.Value = ""
wa.Documento.Value = ""
wa.Nome.Value = ""
wa.Morada.Value = ""
wa.Levantamento.Value = ""

Sheets("Pedido_Cad").Select

Sheets("BD_Pedidos").Select

    ActiveWorkbook.Worksheets("BD_Pedidos").ListObjects("Pedidos").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("BD_Pedidos").ListObjects("Pedidos").Sort.SortFields. _
        Add2 Key:=Range("Pedidos[[#All],[Nº Pedido]]"), SortOn:=xlSortOnValues, _
        Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BD_Pedidos").ListObjects("Pedidos").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With

Resp = MsgBox("Permanecer na tela de pedido para imprimir o cupom? ", vbYesNo, "Atenção")
    
Salva_Plan
    
If Resp = vbYes Then
    
    ped = converterNumero(ped)
    
    Sheets("Pedido_Novo").Select
    
    Carregar_Alteracao ped, True
    
Application.ScreenUpdating = True
Exit Sub
End If

Sheets("Pedido_Cad").Select
Call CRUD_Read
Exit Sub


Call DesbloquearAbasComSenha
Application.ScreenUpdating = True
SemItens: MsgBox "Você não incluiu nenhum item...", vbCritical
End Sub

Sub Novo_Pedido()
    
    Dim ws As Worksheet
    Dim NPedido As Variant
    
    Sheets("Pedido_Novo").Select
    
    Objetos = GetListboxVariables()
    
    Set ws = ThisWorkbook.Worksheets("BD_Pedidos")
    Set Tabela_Servicos = Objetos(1) 'Range da tabela
    Set wa = Objetos(6)

    Set Resumo_Servicos = Objetos(3)
    On Error Resume Next
    Sheets("Resumo_Pedido").Range("TB_Resumo").Delete
    On Error GoTo 0
    
    ActiveSheet.Shapes("Picture 10288").Visible = False
    
    Last = Last_Item(ws)
    
    NPedido = ws.Cells(Last, 1) + 1
    
    If NPedido = 1 Then
        NPedido = InputBox("informe o numero inicial de pedido")
        If NPedido = "" Then
            NPedido = 1
        End If
    End If
    
    
    
    wa.N_Pedido.Value = Format(NPedido, "00000")
    
    wa.Lista_Pedido.Clear
    
    wa.Data.Value = ""
    wa.Documento.Value = ""
    wa.Nome.Value = ""
    wa.Morada.Value = ""
    wa.Levantamento.Value = ""
    wa.Telefone.Value = ""
    
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10254")).Visible = True
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10262")).Visible = False
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10263")).Visible = False


End Sub

Sub Carregar_Alteracao(Linha, Optional Acp As Boolean)
'SpeedOn
Application.ScreenUpdating = False
Dim wl As Worksheet
Dim wr As Worksheet



Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)


Set wl = ThisWorkbook.Worksheets("BD_Pedidos")
Set wr = ThisWorkbook.Worksheets("Resumo_Pedido")
Set wn = ThisWorkbook.Worksheets("Pedido_Novo")
On Error Resume Next
wr.Range("TB_Resumo").Delete

If Acp = False Then
    Pedido = CInt(RemoveLeadingZeros(Resumo_Servicos.List(Linha, 0)))
Else
    Pedido = CInt(Linha)
End If
If Pedido = "" Then
    Exit Sub
End If

On Error GoTo 0
To_End = Last_Item(wl)

repeticoes = WorksheetFunction.CountIf(Planilha16.Range("A:A"), Pedido)
Rept = 0
X = 2
validador = False
For i = 2 To To_End
    
    N_BD_Pedido = wl.Cells(i, 1)
    
    If Pedido = N_BD_Pedido Then
        If dados_principais = "" Then
            dados_principais = 1
        End If
        If dados_principais = 1 Then
                Sheets("Pedido_Novo").Select
                
                ActiveSheet.Shapes("Picture 10288").Visible = True
                
                wn.N_Pedido.Value = Format(wl.Cells(i, 1), "00000")
                wn.Data.Value = wl.Cells(i, 2)
                wn.Documento.Value = wl.Cells(i, 4)
                wn.Nome.Value = wl.Cells(i, 5)
                wn.Morada.Value = wl.Cells(i, 6)
                wn.Levantamento.Value = wl.Cells(i, 3)
                wn.Pagamento.Value = wl.Cells(i, 16)
                wn.Modelo.Value = wl.Cells(i, 17)
                
                Planilha14.Cells(1, 22) = wl.Cells(i, 16)
                
                Planilha9.Cells(43, 6) = wl.Cells(i, 22)
                Planilha9.Cells(43, 7) = wl.Cells(i, 23)
                Planilha9.Cells(43, 8) = wl.Cells(i, 24)
                dados_principais = 2
        End If
        
            For c = 1 To 6
                wr.Cells(X, c) = wl.Cells(i, c + 6)
                'validador = True
                If c = 6 Then
                    X = X + 1
                End If
                
            Next c
            Rept = Rept + 1
    End If
    
    c = c + 1
If Rept = repeticoes Then
    i = To_End
End If
Next i

backup_pagamento = wn.Pagamento.Value
backup_modelo = wn.Modelo.Value

    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10254")).Visible = False
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10262")).Visible = True
    ActiveSheet.Shapes.Range(Array("Rounded Rectangle 10263")).Visible = True


Objetos = GetListboxVariables()

Set wa = Objetos(6)
Set Resumo_Servicos = Objetos(3) 'Lista de serviços

Set ws = ThisWorkbook.Worksheets("Resumo_Pedido")
Set Tb_Resumo_ = ws.Range("TB_Resumo")

Resumo_Servicos.Clear
X = 0
For i = 1 To Tb_Resumo_.Rows.Count
    With Resumo_Servicos
        .AddItem Tb_Resumo_.Cells(i, 1)
        .List(X, 2) = Tb_Resumo_.Cells(i, 2)
        .List(X, 3) = FormatEuro(Tb_Resumo_.Cells(i, 3))
        .List(X, 4) = Tb_Resumo_.Cells(i, 4)
        .List(X, 5) = FormatEuro(Tb_Resumo_.Cells(i, 5))
    End With
    X = X + 1
Next i


Salva_Plan

Call Cumpom_Fiscal
Application.ScreenUpdating = False

wn.Pagamento.Value = backup_pagamento
wn.Modelo.Value = backup_modelo

Carrega_Cliente
Application.ScreenUpdating = True
SpeedOff
Exit Sub
Erro22: Toast "Atenção", "Selecione um dado valido!", 3
Application.ScreenUpdating = True
SpeedOff
End Sub

Sub AlterarDados(Optional Acao As Boolean)



Dim wr As Worksheet
Dim wb As Worksheet
Dim Resp As Variant
Dim tipo As Boolean

Objetos = GetListboxVariables()
    Set wa = Objetos(6)

If Acao = False Then

Resp = MsgBox("Clique em 'Sim' para confirmar a alteração do pedido: " & wa.N_Pedido, vbYesNo, "Atenção")

If Resp = vbNo Then
    Exit Sub
End If
End If

If Acao = True Then
    Resp = MsgBox("Clique em 'Sim' para confirmar a exclusão do pedido: " & wa.N_Pedido, vbYesNo, "Atenção")
        If Resp = vbNo Then
        Exit Sub
    End If
End If

Set wr = ThisWorkbook.Worksheets("Resumo_Pedido")
Set wb = ThisWorkbook.Worksheets("BD_Pedidos")
Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)


Dim wr_End As Variant ' Linha final do resumo
Dim wb_End As Variant ' Linha final do BD
Dim N_Itens As Variant

Pedido = CInt(RemoveLeadingZeros(wa.N_Pedido))

wr_End = Last_Item(wr) - 1
wb_End = Last_Item(wb)

With wb

    For i = 2 To wb_End
    
        If .Cells(i, 1) = Pedido Then
            
            If .Range("Pedidos").ListObject.ListColumns("Nº Pedido").DataBodyRange(i - 1) = Pedido Then
                .Range("Pedidos").ListObject.ListRows(i - 1).Delete
                i = i - 1
            End If
            
        End If
    
    Next i

End With

tipo = True
If Acao = False Then
    Call Salva_Pedido(tipo)
    
Else
    Toast_Info "Pedido: " & wa.N_Pedido & " excluído com sucesso!"
End If



End Sub

Sub Deleta_Pedido_BD()

AlterarDados True

Sheets("Pedido_Cad").Select

Call CRUD_Read

End Sub

Sub Cumpom_Fiscal()

Application.ScreenUpdating = False

Dim wr As Worksheet


Set wr = ThisWorkbook.Worksheets("Resumo_Pedido")
Set wc = ThisWorkbook.Worksheets("Cupom")

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

Total = Range("AE17")

TT_ = Last_Item(wr) - 1

X = 18
i = 18
Contador = 1

    Sheets("Cupom").Select
    
    Do
    
        If wc.Cells(X, 3) <> "" Then
        
            Rows(X & ":" & X).Select
            Selection.Delete Shift:=xlUp
        
        End If
        
    Loop Until wc.Cells(X, 3) = ""
    
    Range("D11") = "Nº " & wa.N_Pedido
    Range("B9") = wa.Nome
    Range("C7") = wa.Telefone
    'Range("C22") = Planilha14.Cells(1, 22)
    'Range("C23") = Planilha14.Cells(1, 23)
    Range("B12") = wa.Modelo
    'Range("C27") = wa.Documento
      
    If TT_ > Contador Then
    
        Do
        
            Rows(i & ":" & i).Select
            Selection.Insert Shift:=xlDown
            Contador = Contador + 1
            
        Loop Until Contador = TT_
        
    End If
    i = 17
    Final = i + (TT_ - 1)
    
    For i = 17 To Final
    
        Cells(i, 1) = wr.Cells(i - 15, 1)
        Cells(i, 3) = wr.Cells(i - 15, 4)
        Cells(i, 4) = wr.Cells(i - 15, 3)
        Cells(i, 5) = wr.Cells(i - 15, 5)
        
    Next i
    
    wa.Select
       
Application.ScreenUpdating = True

End Sub

Sub Carrega_Cliente()

Dim wc As Worksheet

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)


Set wc = ActiveWorkbook.Worksheets("Clientes")

LL = wc.Range("Clientes").Rows.Count

For i = 1 To LL
    
    Telefone = CStr(wc.Range("Clientes").Cells(i, 3).Value)
    Telefone_ = wa.Telefone.Value
    Nome = CStr(wc.Range("Clientes").Cells(i, 2).Value)
    Nome_ = wa.Telefone.Value
    
    If Telefone = Telefone_ And Nome = Nome_ Then
    
        wa.Nome.Value = wc.Range("Clientes").Cells(i, 2).Value
        wa.Morada.Value = wc.Range("Clientes").Cells(i, 4).Value
        wa.Telefone.Value = wc.Range("Clientes").Cells(i, 3).Value
    
    End If

Next i
    
End Sub

Sub Atualiza_Clentes()

Dim wc As Worksheet

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)


Set wc = ActiveWorkbook.Worksheets("Clientes")

LL = wc.Range("Clientes").Rows.Count

For i = 1 To LL
    
    Telefone = CStr(wc.Range("Clientes").Cells(i, 3).Value)
    Telefone_ = wa.Telefone.Value
    Nome = CStr(wc.Range("Clientes").Cells(i, 2).Value)
    Nome_ = wa.Telefone.Value
    
    If Telefone = Telefone_ And Nome = Nome_ Then
        
        wc.Range("Clientes").Cells(i, 1).Value = wa.Documento.Value
        wc.Range("Clientes").Cells(i, 2).Value = wa.Nome.Value
        wc.Range("Clientes").Cells(i, 3).Value = wa.Telefone.Value
        wc.Range("Clientes").Cells(i, 4).Value = wa.Morada.Value
        
        Exit Sub
        
    End If

Next i

Dim Tabela As ListObject
Set Tabela = wc.ListObjects("Clientes")

Dim novaLinha As ListRow
Set novaLinha = Tabela.ListRows.Add

novaLinha.Range(1, 1).Value = wa.Documento.Value
novaLinha.Range(1, 2).Value = wa.Nome.Value
novaLinha.Range(1, 3).Value = wa.Telefone.Value
novaLinha.Range(1, 4).Value = wa.Morada.Value

End Sub

Sub Identificador()

Application.ScreenUpdating = False

Dim wr As Worksheet


Set wr = ThisWorkbook.Worksheets("Resumo_Pedido")
Set wc = ThisWorkbook.Worksheets("Etiqueta")

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set Tabela_Servicos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)
Dim i As Integer
Dim TT As Integer
wc.Select

    Range("A1").Select
    Selection.Copy
    Range("A6:E2707").Select
    ActiveSheet.Paste

Range("A2") = Planilha14.Cells(9, 2)

Range("A3") = Planilha14.Cells(11, 4)
Data = Planilha18.Cells(1, 13)
Range("A4") = Data


TT = Range("H1")

Range("A2:B4").Select

Selection.Copy

If TT > 1 Then
    TT = TT - 1
End If

For i = 1 To TT
    'Copia (ActiveCell.Address)
    resultado = ParOuImpar(i)
    If resultado = "Ímpar" Then
        ActiveCell.Offset(0, 2).Select
        ActiveSheet.Paste
    Else
        ActiveCell.Offset(4, -3).Select
        ActiveSheet.Paste
    End If
Next i

Sheets("Pedido_Novo").Select

End Sub

Function Copia(UltimaPosi)



Range("C38:D40").Select

Selection.Copy



End Function

