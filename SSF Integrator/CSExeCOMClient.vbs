SET obj = CreateObject("SSFIntegration.SSFIntegrationCOMObject")

Private WithEvents in_process As SSFIntegration.SSFIntegrationCOMObject

if NULL <> obj then
    MsgBox "SSFIntegration.SSFIntegrationCOMObject object is created"
Endif


Private sub in_process_RecordRequested(string clientId,string contractId)
    MsgBox.Show(clientId & ": " &  contractId)
end Sub

VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   Caption         =   "SSF Application"
   ClientHeight    =   5235
   ClientLeft      =   9390
   ClientTop       =   645
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   5745
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   5535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "                SSF Client"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents in_process As InProcessClass
Attribute in_process.VB_VarHelpID = -1
Private connector As ConnectorCl

Private Sub in_process_EventSSF(ScreenPopData As Collection, line_number As Integer, input_EventInfo As DesktopToolkitX.TEventInfo)

    Form1.Display_Attach_Data ScreenPopData

End Sub

Private Sub Form_Load()
    Set connector = New ConnectorCl
    Set in_process = connector.InProcessClass
    in_process.Increment_Client_Counter
End Sub

Private Sub Form_Unload(Cancel As Integer)
    in_process.Server_Shut_Down
End Sub

Public Sub Display_Attach_Data(ScreenPopData As Collection)
'displays the attach data within the collection object
'that was send to this client

Set temp_collection = New Collection
List1.Clear
List1.AddItem "caller_id: " & ScreenPopData.Item("caller_id")
List1.AddItem "dialed_number: " & ScreenPopData.Item("dialed_number")
List1.AddItem "high_value_customer: " & ScreenPopData.Item("high_value_customer")
List1.AddItem "client_name: " & ScreenPopData.Item("client_name")
List1.AddItem "address: " & ScreenPopData.Item("address")
List1.AddItem "dob: " & ScreenPopData.Item("dob")
List1.AddItem "tax_id: " & ScreenPopData.Item("tax_id")
List1.AddItem "client_id: " & ScreenPopData.Item("client_id")
List1.AddItem "pin_verified: " & ScreenPopData.Item("pin_verified")
List1.AddItem "csr_verified: " & ScreenPopData.Item("csr_verified")
List1.AddItem "ivr_activity: " & ScreenPopData.Item("ivr_activity")
List1.AddItem "ivr_exit_point: " & ScreenPopData.Item("ivr_exit_point")
List1.AddItem "option_out: " & ScreenPopData.Item("option_out")
List1.AddItem "notes: " & ScreenPopData.Item("notes")

Set temp_collection = ScreenPopData.Item("policy_information")
Dim policy_count, i As Integer
policy_count = temp_collection.Count
For i = 1 To policy_count
    List1.AddItem "policyinfo_" & Str(i) & ": " & temp_collection.Item(i)
Next i

End Sub




