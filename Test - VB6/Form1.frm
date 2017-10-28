VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2625
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2445
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents obj_a As SSFIntegration.SSFComMessageProxy
Attribute obj_a.VB_VarHelpID = -1

Public Sub obj_a_Requested(ByVal clientId As String, ByVal taxId As String, ByVal contractId As String, ByVal contractCount As Long)
    MsgBox "obj_a_Requested: got event: clientId=" & clientId & " taxId=" & taxId & " contractId=" & contractId & " contractCount=" + Str(contractCount)
    Text1.Text = Text1.Text & "obj_a_Requested: got event: clientId=" & clientId & " taxId=" & taxId & " contractId=" & contractId & " contractCount=" + Str(contractCount) & vbNewLine
End Sub

Private Sub Form_Load()
        Set obj_a = New SSFComMessageProxy
End Sub


Sub Main()
    MsgBox "in sub main"
End Sub

