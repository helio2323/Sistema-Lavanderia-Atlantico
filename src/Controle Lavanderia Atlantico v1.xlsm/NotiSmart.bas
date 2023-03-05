Attribute VB_Name = "NotiSmart"
Function Notify_SmartPho(Planilha As String, Msg As String)

    'ativar "Microsoft XML V6.0 em ferramentas > Referências"

    Dim pb_title As String
    Dim pb_title_input As String
    Dim pb_body As String
    Dim pb_body_input As String
    Dim ACCESS_TOKEN As String
    Dim Url As String
    Dim postData As String
    Dim Request As Object
    
    '=======================================
    'CHANGE THE FOLLOWING
    ACCESS_TOKEN_INPUT = "o.EWzJbfjPQeGJnK7583UPjSNHl53YyZrS" 'token gerado no site
    pb_title_input = Planilha 'titulo da mensagem
    pb_body_input = Msg  'corpo da mensagem"
    '=======================================

    'Authentication
    ACCESS_TOKEN = "Bearer " & ACCESS_TOKEN_INPUT
    
    'Variables
    pb_title = """" & pb_title_input & """"
    pb_body = """" & pb_body_input & """"

    'Use XML HTTP
    Set Request = CreateObject("MSXML2.XMLHTTP")

    'Specify Target URL
    Url = "https://api.pushbullet.com/v2/pushes"

    'Open Post Request
    Request.Open "Post", Url, False

    'Request Header
    Request.setRequestHeader "Authorization", ACCESS_TOKEN
    Request.setRequestHeader "Content-Type", "application/json;charset=UTF-8"

    'Concatenate PostData
    postData = "{""type"":""note"",""title"":" & pb_title & ",""body"":" & pb_body & "}"

    'Send the Post Data
    Request.send postData

    '[OPTIONAL] Get response text (result)
    'MsgBox Request.responseText
    
End Function


Sub teste()

Notify_SmartPho "OI tudo bem"

End Sub

