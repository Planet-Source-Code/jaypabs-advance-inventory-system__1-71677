VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPayment 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Payment"
   ClientHeight    =   4380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4380
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame PaymentOption 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Option"
      Height          =   2265
      Left            =   240
      TabIndex        =   3
      Top             =   1350
      Width           =   2895
      Begin VB.ComboBox cbCA 
         Height          =   315
         ItemData        =   "frmPayment.frx":0000
         Left            =   180
         List            =   "frmPayment.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   720
         Width           =   2565
      End
      Begin VB.ComboBox cbPT 
         Height          =   315
         ItemData        =   "frmPayment.frx":001C
         Left            =   180
         List            =   "frmPayment.frx":0026
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1470
         Width           =   2565
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   240
         Index           =   12
         Left            =   180
         TabIndex        =   17
         Top             =   1140
         Width           =   2190
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Charge Account"
         Height          =   240
         Index           =   6
         Left            =   180
         TabIndex        =   16
         Top             =   420
         Width           =   2190
      End
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   345
      Left            =   6240
      TabIndex        =   1
      Top             =   3840
      Width           =   1005
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   13
      Top             =   720
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   53
   End
   Begin VB.Frame Check 
      Caption         =   "Check"
      Height          =   1065
      Left            =   3360
      TabIndex        =   5
      Top             =   2550
      Visible         =   0   'False
      Width           =   3825
      Begin MSDataListLib.DataCombo dcBank 
         Height          =   315
         Left            =   1650
         TabIndex        =   12
         Top             =   570
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtCheckNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1650
         TabIndex        =   6
         Top             =   210
         Width           =   1935
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         Height          =   255
         Left            =   300
         TabIndex        =   11
         Top             =   630
         Width           =   1245
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number:"
         Height          =   255
         Left            =   300
         TabIndex        =   7
         Top             =   270
         Width           =   1245
      End
   End
   Begin VB.Frame Credit 
      Caption         =   "Credit"
      Height          =   1095
      Left            =   3360
      TabIndex        =   4
      Top             =   1350
      Visible         =   0   'False
      Width           =   3825
      Begin VB.ComboBox cbBI 
         Height          =   315
         ItemData        =   "frmPayment.frx":0037
         Left            =   1215
         List            =   "frmPayment.frx":0041
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   240
         Width           =   2565
      End
      Begin VB.TextBox txtDP 
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1215
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   615
         Width           =   1500
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Billed In   Full Payment"
         Height          =   240
         Index           =   7
         Left            =   30
         TabIndex        =   21
         Top             =   270
         Width           =   2145
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Down Payment"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   -885
         TabIndex        =   20
         Top             =   615
         Width           =   2040
      End
   End
   Begin VB.Frame Cash 
      Caption         =   "Cash"
      Height          =   735
      Left            =   3360
      TabIndex        =   8
      Top             =   1350
      Width           =   3825
      Begin VB.TextBox txtCashAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   225
         Left            =   510
         TabIndex        =   10
         Top             =   270
         Width           =   915
      End
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000004&
      Height          =   495
      Left            =   90
      Top             =   3780
      Width           =   7305
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   3675
      Left            =   90
      Top             =   90
      Width           =   7305
   End
   Begin VB.Label lblCustomer 
      BackColor       =   &H80000009&
      Caption         =   "Customer:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   300
      TabIndex        =   2
      Top             =   210
      Width           =   6915
   End
   Begin VB.Label lblTotalAmount 
      BackColor       =   &H80000009&
      Caption         =   "Receipt Total:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   300
      TabIndex        =   0
      Top             =   870
      Width           =   6885
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK               As Integer
Public ReceiptBatchPK  As Integer
Public ClientID         As Integer
Public strCustomer      As String
Public strRefNo         As String
Public TotalAmount      As Currency

Dim AmountPaid          As Currency
Dim intPaymentOption    As Integer

Private Sub cbBI_Click()
    'Not paid Option
    If cbBI.ListIndex = 0 Then
        txtDP.Enabled = False
        txtDP.Text = "0.00"
        cbPT.ListIndex = -1
        cbPT.Enabled = False
        
        Check.Visible = False
    Else 'If Partial
        txtDP.Enabled = True
        cbPT.ListIndex = 0
        cbPT.Enabled = True
    End If
End Sub

Private Sub cbCA_Click()
    txtDP.Text = "0.00"
    'Charge Account Option
    If cbCA.ListIndex = 1 Then 'If Credit
        cbBI.Visible = True
        Label4.Visible = True
        txtDP.Visible = True
        cbPT.ListIndex = -1
        cbBI.ListIndex = 0
        cbPT.Enabled = False
        
        Credit.Visible = True
        Cash.Visible = False
        
    Else 'If Cash
        cbBI.Visible = False
        Label4.Visible = False
        txtDP.Visible = False
        cbPT.ListIndex = 0
        cbPT.Enabled = True
    
        Credit.Visible = False
        Cash.Visible = True
        Check.Visible = False
    End If
End Sub

Private Sub cbPT_Click()
    If cbPT.ListIndex = 0 Then
        Check.Visible = False
    Else
        Check.Visible = True
    End If
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo err
    
    'Save account to Customer's Ledger
    Dim RSClientsLedger As New Recordset
    Dim LedgerID As Integer
    
    RSClientsLedger.CursorLocation = adUseClient
    RSClientsLedger.Open "SELECT * FROM Clients_Ledger WHERE LedgerID=" & 0, CN, adOpenStatic, adLockOptimistic

    If cbCA.Text = "" Then
        MsgBox "Please select charge account.", vbExclamation
        cbCA.SetFocus
        
        Exit Sub
    End If
        
    If cbCA.ListIndex = 0 Then
        If cbPT.Text = "" Then
            MsgBox "Please select Payment Type.", vbExclamation
            cbPT.SetFocus
            
            Exit Sub
        End If
        AmountPaid = toMoney(txtCashAmount.Text)
    Else
        If cbBI.Text = "" Then
            MsgBox "Please select credit type.", vbExclamation
            cbBI.SetFocus
            
            Exit Sub
        End If
        
        If cbBI.ListIndex = 1 Then
            AmountPaid = toMoney(txtDP.Text)
        End If
    End If
    
    With RSClientsLedger
        
        If cbBI.ListIndex <> 0 Then
            .AddNew
            
            LedgerID = getIndex("Clients_Ledger")
    
            !LedgerID = LedgerID
            !ReceiptID = PK
            !ReceiptBatchID = ReceiptBatchPK
            !ClientID = ClientID
            !RefNo = strRefNo
            !ChargeAccount = cbCA.Text
            !PaymentType = cbPT.Text
            !Credit = toMoney(AmountPaid)
            
            If cbPT.ListIndex = 1 Then
                CN.Execute "INSERT INTO Payments_Checks ( LedgerID, CheckNo, BankID, [Date] ) " _
                        & "VALUES (" & LedgerID & ", " & txtCheckNo.Text & ", " & dcBank.BoundText & ", Date())"
            End If
            
            .Update
        End If

    End With

    Set RSClientsLedger = Nothing
    
    Unload Me
    
    Exit Sub
err:
    prompt_err err, Name, "cmdUpdate_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM Banks", "Bank", dcBank, "BankID", True
    
    lblCustomer.Caption = lblCustomer.Caption & " " & strCustomer
    lblTotalAmount.Caption = lblTotalAmount.Caption & " " & toMoney(TotalAmount)
    
    intPaymentOption = 1
    
    txtCashAmount.Text = toMoney(TotalAmount)
    
    cbCA.ListIndex = 0
    cbPT.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPayment = Nothing
End Sub
