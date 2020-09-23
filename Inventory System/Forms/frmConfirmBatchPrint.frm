VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConfirmBatchPrint 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   7035
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   7875
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   7875
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   6510
      TabIndex        =   4
      Top             =   5670
      Width           =   1155
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   345
      Left            =   5280
      TabIndex        =   3
      Top             =   5670
      Width           =   1155
   End
   Begin VB.CommandButton cmdDeselectAll 
      Caption         =   "&De-select All"
      Height          =   345
      Left            =   1500
      TabIndex        =   2
      Top             =   5670
      Width           =   1155
   End
   Begin VB.CommandButton cmdSelectAll 
      Caption         =   "&Select All"
      Height          =   345
      Left            =   210
      TabIndex        =   1
      Top             =   5670
      Width           =   1155
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   5250
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   9260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   3598
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "D.R. No."
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Agent"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Printed"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Receipt ID"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   7740
      Y1              =   6120
      Y2              =   6120
   End
   Begin VB.Label Label2 
      Caption         =   "Note:"
      Height          =   345
      Left            =   150
      TabIndex        =   6
      Top             =   6210
      Width           =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   150
      X2              =   7770
      Y1              =   6090
      Y2              =   6090
   End
   Begin VB.Shape Shape1 
      Height          =   5445
      Left            =   60
      Top             =   90
      Width           =   7725
   End
   Begin VB.Label Label1 
      Caption         =   $"frmConfirmBatchPrint.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   810
      TabIndex        =   5
      Top             =   6210
      Width           =   6945
   End
End
Attribute VB_Name = "frmConfirmBatchPrint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public intRoute     As Integer
Public ddate        As Date

Dim rs              As New Recordset
Dim intTotalRecord  As Integer

Private Sub cmdClose_Click()
    rs.Close
    Set rs = Nothing
    
    Unload Me
End Sub

Private Sub cmdDeselectAll_Click()
    Dim I As Integer

    With lvList
        I = 1
        
        For I = 1 To intTotalRecord
            .ListItems(I).Checked = False
        Next I
    End With
End Sub

Private Sub cmdSelectAll_Click()
    Dim I As Integer

    With lvList
        I = 1
        
        For I = 1 To intTotalRecord
            .ListItems(I).Checked = True
        Next I
    End With
End Sub

Private Sub cmdUpdate_Click()
    Dim I As Integer

    With lvList
        I = 1
        
        For I = 1 To intTotalRecord
            rs.MoveFirst
            rs.Find "ReceiptID =" & .ListItems(I).SubItems(4)
            
            If rs.RecordCount > 0 Then
                rs!Printed = .ListItems(I).Checked
                
                rs.Update
            End If
        Next I
    End With
End Sub

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT Company, RefNo, AgentName, Printed, ReceiptID " _
            & "FROM qry_Receipts " _
            & "WHERE Route = '" & intRoute & "' " _
            & "AND DeliveryDate = #" & ddate & "# " _
            & "AND Printed=False", _
            CN, adOpenStatic, adLockOptimistic
    
    If rs.RecordCount > 0 Then
        FillListView lvList, rs, 9, 0, False, True, "ReceiptID"
        
        intTotalRecord = rs.RecordCount
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmConfirmBatchPrint = Nothing
End Sub
