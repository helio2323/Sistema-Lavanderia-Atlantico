Attribute VB_Name = "mdlCalendário"
Option Explicit

'Vetor que armazena todos os Label de dia do Calendário
Dim Rótulos() As New cCalendário

Function GetCalendário() As Date
        
    Dim lTotalRótulos As Long
    Dim ctrl As control
    Dim frm As frmCalendário
    
    Set frm = New frmCalendário
    
    'Atribui cada um dos Label num elemento do vetor da classe
    For Each ctrl In frm.Controls
        If ctrl.name Like "l?c?" Then
            lTotalRótulos = lTotalRótulos + 1
            ReDim Preserve Rótulos(1 To lTotalRótulos)
            Set Rótulos(lTotalRótulos).lblGrupo = ctrl
        End If
    Next ctrl

    frm.Show
    
    'Se a data escolhida for nula ou inválida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalendário = frm.Tag
    Else
        GetCalendário = Date
    End If
        
    Unload frm

End Function
