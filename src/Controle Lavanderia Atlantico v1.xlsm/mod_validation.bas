Attribute VB_Name = "mod_validation"
Sub import_situacao()

    Dim PesqTb  As QueryTable
    Dim Url     As String
    Dim M As String
    Dim Shet As String
    
    On Error GoTo trata_erro
    Email = Select_Cx("Cx_Email")
    Senha = Planilha1.TextBox1.Value
    
    plan_ativacao.Cells(2, 2).Value = Email
    plan_ativacao.Cells(2, 3).Value = Senha
    
    For Each PesqTb In plan_ativacao.QueryTables
        plan_ativacao.Range("A5:G1000000").Clear
        PesqTb.Delete
    Next
    
    
    Url = "Url;" & "https://docs.google.com/spreadsheets/d/e/2PACX-1vRwhQoQohIGvYuka6fAiaIw-HoQup4jROGI7POumrx3P3ry08V3n2hm5vlKfaSeks4Agfbtnd0HVUqA/pubhtml"
    Set PesqTb = plan_ativacao.QueryTables.Add(Url, plan_ativacao.Range("A5"))
    
    With PesqTb
        .BackgroundQuery = False
        .RefreshOnFileOpen = False
        .name = "validação"
        .WebFormatting = xlWebFormattingAll
        .WebTables = "1"
        .Refresh
    End With
    
    Call validacao
    
    M = "O cliente " & Email & " logou na planilha com sucesso em" & Date & " " & Time
    Shet = ActiveWorkbook.name
    
    Notify_SmartPho Shet, M
    'Criar continuação para fazer o login
    
    ToastGren "Login realizado com sucesso, seu sistema esta carregando..."
    
    Exit Sub
trata_erro:
    
    MsgBox "Ocorreu um erro na validação da planilha!" & vbNewLine & "Verifique sua conexão com a internet", vbCritical, "Organic Sheets"
    'ThisWorkbook.Close SaveChanges:=False
    
End Sub

Sub validacao()
    
    UltimaLinha = plan_ativacao.Cells(Rows.Count, 3).End(xlUp).Row
    
    For i = 7 To UltimaLinha
    
        If plan_ativacao.Cells(i, 3).Value = plan_ativacao.Range("B2") Then
            
            If plan_ativacao.Cells(i, 5) = "ativo" And plan_ativacao.Cells(i, 4) = plan_ativacao.Range("C2") Then
                plan_ativacao.Cells(i, 7).Value = "Logado"
                Planilha6.Select
                Exit Sub
            ElseIf plan_ativacao.Cells(i, 5) = "ativo" And plan_ativacao.Cells(i, 4) <> plan_ativacao.Range("C2") Then
                MsgBox "Você não tem permissão para acessar a planilha desse computador" & vbNewLine & "Entre em contato com o desenvolvedor", vbCritical, "Acesso Negado"
                'ThisWorkbook.Close SaveChanges:=False
            Else
                MsgBox "Você não tem permissão para acessar a planilha" & vbNewLine & "Entre em contato com o desenvolvedor", vbCritical, "Acesso Negado"
                'ThisWorkbook.Close SaveChanges:=False
            End If
        
                MsgBox "Você não tem permissão para acessar a planilha desse computador" & vbNewLine & "Entre em contato com o desenvolvedor", vbCritical, "Acesso Negado"
                'ThisWorkbook.Close SaveChanges:=False
        End If
        
    Next i
    
    
    
End Sub
