VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Acom 
   Caption         =   "Acompanhamento"
   ClientHeight    =   5535
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9720
   OleObjectBlob   =   "Acom.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Acom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Call Atualiza_Status
End Sub

Private Sub CommandButton3_Click()

    Call Alterar_Acompanhamento
    
End Sub

'@Folder("VBAProject")
Private Sub Label1_Click()

End Sub

Private Sub UserForm_Initialize()

SetClassTextBox Me
    
Call Combos
    
Dim y As Long

Pedido = ActiveCell.Value2

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set bd_pedidos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

Acom.L_N_Pedido.Caption = Pedido

Acom.ListBox1.ColumnCount = 5
Acom.ListBox1.ColumnWidths = "100;40;40;40;40"

y = 0
With bd_pedidos
    For i = 1 To bd_pedidos.Rows.Count

        If .Cells(i, 1) = Pedido Then
            
            Acom.Cb_Status.Value = .Cells(i, 12)
            Acom.Cb_Responsavel.Value = .Cells(i, 13)
            Acom.Cb_Pagamento.Value = .Cells(i, 16)
                
            Acom.ListBox1.AddItem (.Cells(i, 6))
            Acom.ListBox1.List(y, 0) = .Cells(i, 7)
            Acom.ListBox1.List(y, 1) = FormatEuro(.Cells(i, 9))
            Acom.ListBox1.List(y, 2) = .Cells(i, 10)
            Acom.ListBox1.List(y, 3) = FormatEuro(.Cells(i, 11))

            y = y + 1
        End If
        
    Next i
End With


    
End Sub



Sub Combos()

Dim wf As Worksheet

Set wf = ActiveWorkbook.Worksheets("Funcionários")

Objetos = GetListboxVariables()

Set ws = Objetos(0)
Set bd_pedidos = Objetos(1) 'Range da tabela
Set Servicos_ComboBox = Objetos(2)
Set Resumo_Servicos = Objetos(3)
Set wa = Objetos(6)

With Acom
    
    .Cb_Status.Clear
    .Cb_Responsavel.Clear
    .Cb_Pagamento.Clear
    
    .Cb_Status.AddItem "Em Andamento"
    .Cb_Status.AddItem "Aguardando Levantamento"
    .Cb_Status.AddItem "Entregue"
    
    For i = 1 To wf.Range("Funcionarios").Rows.Count
    
        .Cb_Responsavel.AddItem wf.Range("Funcionarios").Cells(i, 2)
    
    Next i
    
    .Cb_Pagamento.AddItem "Aguardando Pagamento"
    .Cb_Pagamento.AddItem "Pago"
    
    
End With


End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
