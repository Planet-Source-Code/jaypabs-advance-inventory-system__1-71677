VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmPayVendorBills 
   BackColor       =   &H00808000&
   BorderStyle     =   0  'None
   ClientHeight    =   7230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame PaymentOption 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Payment Option"
      Height          =   2295
      Left            =   210
      TabIndex        =   13
      Top             =   1110
      Width           =   8565
      Begin VB.ComboBox cbPT 
         Height          =   315
         ItemData        =   "frmPayVendorBills.frx":0000
         Left            =   2880
         List            =   "frmPayVendorBills.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   570
         Width           =   1635
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   4620
         TabIndex        =   4
         Top             =   570
         Width           =   1935
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   375
         Left            =   7110
         TabIndex        =   7
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Frame Check 
         Caption         =   "Check"
         Enabled         =   0   'False
         Height          =   1065
         Left            =   210
         TabIndex        =   14
         Top             =   1050
         Width           =   3825
         Begin VB.TextBox txtCheckNo 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   1650
            TabIndex        =   5
            Top             =   210
            Width           =   1935
         End
         Begin ctrlNSDataCombo.NSDataCombo nsdBank 
            Height          =   315
            Left            =   1650
            TabIndex        =   6
            Top             =   570
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   556
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number:"
            Height          =   255
            Left            =   300
            TabIndex        =   16
            Top             =   270
            Width           =   1245
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bank:"
            Height          =   255
            Left            =   300
            TabIndex        =   15
            Top             =   630
            Width           =   1245
         End
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdRefNo 
         Height          =   315
         Left            =   240
         TabIndex        =   2
         Top             =   570
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   240
         Index           =   6
         Left            =   2910
         TabIndex        =   19
         Top             =   270
         Width           =   1560
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   225
         Left            =   4650
         TabIndex        =   18
         Top             =   300
         Width           =   915
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ref No."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   270
         Width           =   1260
      End
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "&Update"
      Height          =   345
      Left            =   6900
      TabIndex        =   10
      Top             =   6690
      Width           =   1005
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   7950
      TabIndex        =   11
      Top             =   6690
      Width           =   1005
   End
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   1050
      TabIndex        =   0
      Top             =   690
      Width           =   2145
   End
   Begin VB.TextBox txtRemarks 
      Height          =   915
      Left            =   1020
      MultiLine       =   -1  'True
      TabIndex        =   8
      Top             =   5460
      Width           =   2835
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmPayVendorBills.frx":001B
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Remove"
      Top             =   3630
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.TextBox txtGross 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7230
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   5340
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   4020
      TabIndex        =   1
      Top             =   690
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   556
      _Version        =   393216
      Format          =   20643841
      CurrentDate     =   39250
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1680
      Left            =   240
      TabIndex        =   20
      Top             =   3570
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2963
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   1091552
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H80000004&
      Height          =   495
      Left            =   90
      Top             =   6630
      Width           =   8955
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000004&
      Height          =   6465
      Left            =   90
      Top             =   90
      Width           =   8955
   End
   Begin VB.Label lblCustomer 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Supplier:"
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
      Left            =   240
      TabIndex        =   25
      Top             =   180
      Width           =   6915
   End
   Begin VB.Label Labels 
      BackStyle       =   0  'Transparent
      Caption         =   "Ref No.:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   300
      TabIndex        =   24
      Top             =   720
      Width           =   690
   End
   Begin VB.Label Label2 
      Caption         =   "Remarks"
      Height          =   315
      Left            =   210
      TabIndex        =   23
      Top             =   5460
      Width           =   765
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Left            =   3420
      TabIndex        =   22
      Top             =   720
      Width           =   555
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Total:"
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
      Left            =   6030
      TabIndex        =   21
      Top             =   5340
      Width           =   1125
   End
End
Attribute VB_Name = "frmPayVendorBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PaymentPK           As Integer
Public PK               As String
Public strRefNo         As String
Public strVendor        As String
Public ForwarderPK       As Long
Public blnNew           As Boolean

Public TotalAmount      As Currency
Dim AmountPaid          As Currency
Dim intPaymentOption    As Integer

Dim cGross              As Currency 'Gross Amount
Dim cRowCount           As Integer

Private Sub btnAdd_Click()
    If nsdRefNo.Text = "" Or cbPT.Text = "" Or txtAmount.Text = "" Then nsdRefNo.SetFocus: Exit Sub

    If AmountPaid > TotalAmount Then
        MsgBox "Amount to be paid must be less than the total amount due.", vbInformation
        Exit Sub
    End If

    If cbPT.ListIndex = 1 And nsdBank.BoundText = "" Then 'if check
        MsgBox "Please don't leave Check No. or Bank field blank", vbExclamation
        
        txtCheckNo.SetFocus
        
        Exit Sub
    End If
    
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 1, nsdRefNo.Text)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 1) = "" Then
                .TextMatrix(1, 1) = nsdRefNo.Text
                .TextMatrix(1, 2) = cbPT.Text
                .TextMatrix(1, 3) = txtAmount.Text
                .TextMatrix(1, 7) = IIf(nsdRefNo.BoundText = "", nsdRefNo.Tag, nsdRefNo.BoundText)
                
                If cbPT.ListIndex = 1 Then
                    .TextMatrix(1, 4) = txtCheckNo.Text
                    .TextMatrix(1, 5) = nsdBank.Text
                    .TextMatrix(1, 6) = nsdBank.BoundText
                End If
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdRefNo.Text
                .TextMatrix(.Rows - 1, 2) = cbPT.Text
                .TextMatrix(.Rows - 1, 3) = txtAmount.Text
                .TextMatrix(.Rows - 1, 7) = IIf(nsdRefNo.BoundText = "", nsdRefNo.Tag, nsdRefNo.BoundText)

                If cbPT.ListIndex = 1 Then
                    .TextMatrix(.Rows - 1, 4) = txtCheckNo.Text
                    .TextMatrix(.Rows - 1, 5) = nsdBank.Text
                    .TextMatrix(.Rows - 1, 6) = nsdBank.BoundText
                End If
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cRowCount = cRowCount + 1
        Else
            If MsgBox("Ref No already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow

                'Restore back the invoice amount and discount
                cGross = cGross - toNumber(Grid.TextMatrix(.RowSel, 3))

                .TextMatrix(CurrRow, 1) = nsdRefNo.Text
                .TextMatrix(CurrRow, 2) = cbPT.Text
                .TextMatrix(CurrRow, 3) = txtAmount.Text
                .TextMatrix(CurrRow, 6) = nsdBank.BoundText
                .TextMatrix(CurrRow, 7) = IIf(nsdRefNo.BoundText = "", nsdRefNo.Tag, nsdRefNo.BoundText)
            
                If cbPT.ListIndex = 1 Then
                    .TextMatrix(CurrRow, 4) = txtCheckNo.Text
                    .TextMatrix(CurrRow, 5) = nsdBank.Text
                End If
            Else
                Exit Sub
            End If
        End If
        
        'Add the amount to current load amount
        cGross = cGross + toNumber(txtAmount.Text)
        txtGross.Text = Format$(cGross, "#,##0.00")
        
        'Highlight the current row's column
        .ColSel = 5
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update gross to current purchase amount
        cGross = cGross - toNumber(Grid.TextMatrix(.RowSel, 3))
        txtGross.Text = Format$(cGross, "#,##0.00")
        
        'Update the record count
        cRowCount = cRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cbPT_Click()
    If cbPT.ListIndex = 0 Then
        Check.Enabled = False
    Else
        Check.Enabled = True
    End If
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdUpdate_Click()
On Error GoTo err

    'Validate entry
    If txtRefNo.Text = "" Then
        MsgBox "Please enter Ref No.", vbInformation
        txtRefNo.SetFocus
        Exit Sub
    End If
         
    If cRowCount < 1 Then
        MsgBox "Please enter bill(s) to pay.", vbExclamation
        nsdRefNo.SetFocus
        Exit Sub
    End If
         
    PaymentPK = getIndex("Vendors_Payments")
   
    'Save account to Customer's Ledger
    Dim RSPayments As New Recordset

    RSPayments.CursorLocation = adUseClient
    RSPayments.Open "SELECT * FROM Vendors_Payments WHERE PaymentID=" & 0, CN, adOpenStatic, adLockOptimistic
    
    CN.BeginTrans
    
    With RSPayments
        .AddNew
        
        !PaymentID = PaymentPK
        !VendorID = PK
        !RefNo = txtRefNo.Text
        !Date = dtpDate.Value
        !Remarks = txtRemarks.Text
        
        .Update
    End With
    
    'Save account to Customer's Ledger
    Dim RSLedger As New Recordset

    RSLedger.CursorLocation = adUseClient
    RSLedger.Open "SELECT * FROM Vendors_Ledger WHERE LedgerID=" & 0, CN, adOpenStatic, adLockOptimistic

    DeleteItems

    Dim c As Integer

    With Grid
        'Save the details of the records
        For c = 1 To cRowCount
            .Row = c
 
            RSLedger.AddNew
            
            RSLedger!VendorID = PK
            RSLedger!POID = .TextMatrix(c, 7)
            RSLedger!PaymentID = PaymentPK
            RSLedger!RefNo = .TextMatrix(c, 1)
            RSLedger!Date = dtpDate.Value
            RSLedger!PaymentType = .TextMatrix(c, 2)
            RSLedger!Credit = .TextMatrix(c, 3)
            
            If .TextMatrix(c, 2) = "Check" Then
                RSLedger!CheckNo = .TextMatrix(c, 4)
                RSLedger!BankID = .TextMatrix(c, 6)
            End If
            
            RSLedger!AddedByFK = CurrUser.USER_PK
            
            RSLedger.Update
        Next c
    End With

    CN.CommitTrans
        
    Set RSPayments = Nothing
    Set RSLedger = Nothing
    
    frmVendorBalance.RefreshRecords
    
    Unload Me
    
    Exit Sub
err:
    CN.RollbackTrans
    
    prompt_err err, Name, "cmdUpdate_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    InitGrid
    InitNSD
    
'    bind_dc "SELECT * FROM Banks", "Bank", dcBank, "BankID", True
    
    lblCustomer.Caption = lblCustomer.Caption & " " & strVendor
    
    intPaymentOption = 1
    
    dtpDate.Value = Date
    
    cbPT.ListIndex = 0
    txtAmount.Text = toMoney(TotalAmount)
    nsdRefNo.Text = strRefNo
    nsdRefNo.Tag = ForwarderPK
    
    blnNew = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmPayVendorBills = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If Grid.Rows = 2 And Grid.TextMatrix(1, 1) = "" Then '1 = RefNo
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub nsdRefNo_Change()
    If blnNew = True Then Exit Sub
    
    txtAmount.Text = toMoney(nsdRefNo.getSelValueAt(3))

    TotalAmount = toMoney(txtAmount.Text)
End Sub

Private Sub txtAmount_Change()
    AmountPaid = toNumber(txtAmount.Text)
End Sub

Private Sub txtAmount_GotFocus()
    HLText txtAmount
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtAmount_Validate(Cancel As Boolean)
    txtAmount.Text = toMoney(txtAmount.Text)
End Sub

Private Sub InitNSD()
    'For Ref No
    With nsdRefNo
        .ClearColumn
        .AddColumn "Ref No", 1300.89
        .AddColumn "Company", 1994.89
        .AddColumn "Balance", 1264.88
        
        .Connection = CN.ConnectionString
        .sqlFields = "RefNo,Company,Balance,DueDate,POID"
        .sqlTables = "qry_Vendor_Balance_Per_Forwarder"
        .sqlwCondition = "VendorID=" & PK
        .sqlSortOrder = "DueDate ASC"
        
        .BoundField = "POID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Unpaid Invoices"
    End With

    'For Bank
    With nsdBank
        .ClearColumn
        .AddColumn "Bank", 1300.89
        .AddColumn "Branch", 1994.89
        
        .Connection = CN.ConnectionString
        .sqlFields = "Bank,Branch,BankID"
        .sqlTables = "Banks"
        .sqlSortOrder = "Bank ASC"
        
        .BoundField = "BankID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Banks"
    End With
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 8
        .ColSel = 7
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1400
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColWidth(4) = 1300
        .ColWidth(5) = 1300
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Ref No"
        .TextMatrix(0, 2) = "Payment Type"
        .TextMatrix(0, 3) = "Amount"
        .TextMatrix(0, 4) = "Check No"
        .TextMatrix(0, 5) = "Bank"
        .TextMatrix(0, 6) = "BankID"
        .TextMatrix(0, 7) = "POID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbRightJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSLedger As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSLedger.CursorLocation = adUseClient
    RSLedger.Open "SELECT * FROM Vendors_Ledger WHERE PaymentID='" & PaymentPK & "'", CN, adOpenStatic, adLockOptimistic
    If RSLedger.RecordCount > 0 Then
        RSLedger.MoveFirst
        While Not RSLedger.EOF
            CurrRow = getFlexPos(Grid, 1, RSLedger!RefNo)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Vendors_Ledger", "LedgerID", "", True, RSLedger!LedgerID
                End If
            End With
            RSLedger.MoveNext
        Wend
    End If
    
    RSLedger.Close
    Set RSLedger = Nothing
End Sub
