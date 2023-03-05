VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Clientes 
   Caption         =   "Lista de Clientes"
   ClientHeight    =   6870
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6825
   OleObjectBlob   =   "Clientes.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Clientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

Linha = Me.ListBox1.ListIndex

Planilha9.Documento = Me.ListBox1.List(Linha, 0)
Planilha9.Nome = Me.ListBox1.List(Linha, 1)
Planilha9.Morada = Me.ListBox1.List(Linha, 3)
Planilha9.Telefone = Me.ListBox1.List(Linha, 2)

Me.Hide

End Sub

Private Sub TextBox1_Change()
Dim Tabela As Range
Dim Pesquisa As String
Dim i As Long
Dim j As Long

Set Tabela = Planilha13.Range("Clientes")
Pesquisa = Me.TextBox1.text

With Me.ListBox1
    .Clear ' limpa a lista antes de preenchê-la novamente
    For i = 0 To Tabela.Rows.Count - 1
        If InStr(1, Tabela.Cells(i + 1, 2), Pesquisa, vbTextCompare) > 0 Then ' verifica se o nome contém a pesquisa
            .AddItem '' adiciona um item em branco
            For j = 1 To Tabela.Columns.Count ' começa a partir da coluna 1
                .List(.ListCount - 1, j - 1) = Tabela.Cells(i + 1, j) ' ajusta a referência da coluna
            Next j
        End If
    Next i
End With

End Sub



Private Sub UserForm_Initialize()
    SetClassTextBox Me
    
Me.TextBox1.Value = ""
    
Dim Tabela As Range
Set Tabela = Planilha13.Range("Clientes")
Dim numColunas As Integer
numColunas = Tabela.Columns.Count

With Me.ListBox1
    .Clear ' limpa a lista antes de preenchê-la novamente
    For i = 0 To Tabela.Rows.Count - 1
        'If InStr(1, Tabela.Cells(i + 1, 2), Pesquisa, vbTextCompare) > 0 Then ' verifica se o nome contém a pesquisa
            .AddItem '' adiciona um item em branco
            For j = 1 To Tabela.Columns.Count ' começa a partir da coluna 1
                .List(.ListCount - 1, j - 1) = Tabela.Cells(i + 1, j) ' ajusta a referência da coluna
            Next j
        'End If
    Next i
End With

    
    
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Unload Me
End Sub
