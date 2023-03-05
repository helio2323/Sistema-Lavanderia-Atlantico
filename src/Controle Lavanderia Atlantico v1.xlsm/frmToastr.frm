VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmToastr 
   Caption         =   "frmToastr"
   ClientHeight    =   1335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8100
   OleObjectBlob   =   "frmToastr.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmToastr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


#If VBA7 Then
Private Declare PtrSafe Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare PtrSafe Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function sndPlaySound Lib "winmm.dll" _
Alias "sndPlaySoundA" (ByVal lpszSoundName As String, _
ByVal uFlags As Long) As Long
  
    
#Else
    Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
    Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
#End If
Public i As Double
Public k As Double
Private Sub lblClose_Click()
Unload Me
End Sub

Private Sub lblMsg_Click()
Unload Me
End Sub

Private Sub UserForm_Activate()

Unload Me
Exit Sub

DONGUDURDUR = False
Me.Height = 97
Me.Width = 413


Me.Left = (Application.Width - 400) / 2
#If VBA7 Then
Dim mWnd As Long

#Else
Dim mWnd As LongPtr

#End If
mWnd = FindWindow(vbNullString, Me.name)

SetWindowRgn mWnd, CreateRoundRectRgn(9, 32, 543, 122, 7, 7), True

Call userform_float
Call progresso
Unload Me
End Sub
Sub userform_float()
Dim j As Double
j = 0
Do While j < 82
DoEvents

j = j + 0.009

Me.Top = j

Loop


End Sub
Public Sub progresso()

Dim i As Double
Dim k As Double
i = 150
Do While i > 1
DoEvents
If i >= 1 Then
    If PORGRESSSTOP = False Then
        i = i - 0.009
        frmToastr.lblProgress.Width = i
         k = i
     End If

End If

Loop

End Sub





