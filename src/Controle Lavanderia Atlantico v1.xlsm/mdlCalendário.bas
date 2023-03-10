Attribute VB_Name = "mdlCalend?rio"
Option Explicit

'Vetor que armazena todos os Label de dia do Calend?rio
Dim R?tulos() As New cCalend?rio

Function GetCalend?rio() As Date
        
    Dim lTotalR?tulos As Long
    Dim ctrl As control
    Dim frm As frmCalend?rio
    
    Set frm = New frmCalend?rio
    
    'Atribui cada um dos Label num elemento do vetor da classe
    For Each ctrl In frm.Controls
        If ctrl.name Like "l?c?" Then
            lTotalR?tulos = lTotalR?tulos + 1
            ReDim Preserve R?tulos(1 To lTotalR?tulos)
            Set R?tulos(lTotalR?tulos).lblGrupo = ctrl
        End If
    Next ctrl

    frm.Show
    
    'Se a data escolhida for nula ou inv?lida, retorna-se a data atual:
    If IsDate(frm.Tag) Then
        GetCalend?rio = frm.Tag
    Else
        GetCalend?rio = Date
    End If
        
    Unload frm

End Function
