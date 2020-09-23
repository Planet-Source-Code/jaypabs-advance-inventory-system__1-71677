VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmBalancePayment 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4365
   ScaleWidth      =   7485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Check 
      Caption         =   "Check"
      Height          =   2055
      Left            =   2910
      TabIndex        =   13
      Top             =   1440
      Visible         =   0   'False
      Width           =   4275
      Begin VB.TextBox txtCheckNo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1890
         TabIndex        =   16
         Top             =   390
         Width           =   1935
      End
      Begin VB.TextBox txtCheckAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1890
         TabIndex        =   14
         Top             =   1170
         Width           =   1935
      End
      Begin MSDataListLib.DataCombo dcBank 
         Height          =   315
         Left            =   1890
         TabIndex        =   15
         Top             =   750
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Check Number:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   450
         Width           =   1665
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Bank:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   810
         Width           =   1665
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Amount:"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1170
         Width           =   1665
      End
   End
   Begin VB.Frame Account 
      Caption         =   "On Account"
      Height          =   1155
      Left            =   2910
      TabIndex        =   6
      Top             =   1440
      Visible         =   0   'False
      Width           =   4275
      Begin VB.TextBox txtAccountAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1710
         TabIndex        =   7
         Top             =   420
         Width           =   1935
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   255
         Left            =   300
         TabIndex        =   8
         Top             =   450
         Width           =   1305
      End
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   345
      Left            =   5190
      TabIndex        =   5
      Top             =   3870
      Width           =   1005
   End
   Begin VB.Frame PaymentOption 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Option"
      Height          =   2235
      Left            =   240
      TabIndex        =   1
      Top             =   1350
      Width           =   2205
      Begin VB.OptionButton Payment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cash"
         Height          =   285
         Index           =   1
         Left            =   300
         TabIndex        =   4
         Top             =   360
         Value           =   -1  'True
         Width           =   1425
      End
      Begin VB.OptionButton Payment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Check"
         Height          =   345
         Index           =   2
         Left            =   300
         TabIndex        =   3
         Top             =   870
         Width           =   1605
      End
      Begin VB.OptionButton Payment 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Account"
         Height          =   405
         Index           =   3
         Left            =   300
         TabIndex        =   2
         Top             =   1380
         Width           =   1605
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   6270
      TabIndex        =   0
      Top             =   3870
      Width           =   1005
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   9
      Top             =   720
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   53
   End
   Begin VB.Frame Cash 
      Caption         =   "Cash"
      Height          =   1515
      Left            =   2910
      TabIndex        =   10
      Top             =   1440
      Width           =   4275
      Begin VB.TextBox txtCashAmount 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1950
         TabIndex        =   11
         Top             =   570
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   225
         Left            =   420
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
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
      TabIndex        =   21
      Top             =   870
      Width           =   6885
   End
   Begin VB.Label lblCustomer 
      BackColor       =   &H80000009&
      Caption         =   "Customer Name:"
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
      TabIndex        =   20
      Top             =   210
      Width           =   6915
   End
   Begin VB.Shape Shape1 
      Height          =   3675
      Left            =   90
      Top             =   90
      Width           =   7305
   End
   Begin VB.Shape Shape2 
      Height          =   495
      Left            =   90
      Top             =   3780
      Width           =   7305
   End
End
Attribute VB_Name = "frmBalancePayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK               As String
Public strCustomer      As String
Public TotalAmount      As Currency

Dim AmountPaid          As Currency
Dim intPaymentOption    As Integer

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo err
    Dim AmountPaid_tmp As Currency
    Dim strRefNo As String
    
    If Payment(1).Value = True Then
        AmountPaid = toMoney(txtCashAmount.Text)
    ElseIf Payment(2).Value = True Then
        AmountPaid = toMoney(txtCheckAmount.Text)
    Else
        AmountPaid = toMoney(txtAccountAmount.Text)
    End If
    
    If AmountPaid > TotalAmount Then
        MsgBox "Overpayment. Customer balance is " & TotalAmount, vbInformation
        Exit Sub
    End If
    
    'Retrieve balance per receipt
    Dim RSBalance As New Recordset

    RSBalance.CursorLocation = adUseClient
    RSBalance.Open "SELECT * FROM qry_Client_Balance_Per_Receipt WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    Do While Not RSBalance.EOF
        strRefNo = RSBalance!RefNo
    
        If AmountPaid > RSBalance!Balance Then
            AmountPaid = AmountPaid - RSBalance!Balance
            Save_Payment RSBalance!Balance, strRefNo
            
            RSBalance.MoveNext
        Else
            Save_Payment AmountPaid, strRefNo
            Exit Do
        End If
    Loop

    Set RSBalance = Nothing
    
    frmCustomerBalance.RefreshRecords
    
    Unload Me
    
    Exit Sub
err:
    prompt_err err, Name, "cmdUpdate_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Save_Payment(Amount As Currency, strRefNo As String)
    'Save account to Customer's Ledger
    Dim RSClientsLedger As New Recordset

    RSClientsLedger.CursorLocation = adUseClient
    RSClientsLedger.Open "SELECT * FROM Clients_Ledger WHERE LedgerID=" & 0, CN, adOpenStatic, adLockOptimistic
    
    With RSClientsLedger
        .AddNew
        
        !ReceiptID = PK
        !ClientID = PK
        !RefNo = strRefNo
        !PaymentOption = intPaymentOption
        !Credit = toMoney(Amount)
        
        If intPaymentOption = 2 Then
            CN.Execute "INSERT INTO Payments_Checks ( LedgerID, CheckNo, BankID, [Date] ) " _
                    & "VALUES (" & PK & ", " & txtCheckNo.Text & ", " & dcBank.BoundText & ", Date())"
        End If
        
        .Update
    End With
   
    Set RSClientsLedger = Nothing
End Sub

Private Sub Form_Load()
    bind_dc "SELECT * FROM Banks", "Bank", dcBank, "BankID", True
    
    lblCustomer.Caption = lblCustomer.Caption & " " & strCustomer
    lblTotalAmount.Caption = lblTotalAmount.Caption & " " & toMoney(TotalAmount)
    
    intPaymentOption = 1
    
    txtCashAmount.Text = toMoney(TotalAmount)
    txtCheckAmount.Text = toMoney(TotalAmount)
    txtAccountAmount.Text = toMoney(TotalAmount)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPayment = Nothing
End Sub

Private Sub Payment_Click(Index As Integer)
    intPaymentOption = Index
    
    If Index = 1 Then
        Cash.Visible = True
        Check.Visible = False
        Account.Visible = False
    ElseIf Index = 2 Then
        Cash.Visible = False
        Check.Visible = True
        Account.Visible = False
    Else
        Cash.Visible = False
        Check.Visible = False
        Account.Visible = True
    End If
End Sub

