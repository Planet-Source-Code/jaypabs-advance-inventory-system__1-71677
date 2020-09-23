VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmCollectionAE 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCollectionAE.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExpenses 
      Caption         =   "&Expenses"
      Height          =   315
      Left            =   6780
      TabIndex        =   45
      Top             =   7650
      Width           =   1185
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmCollectionAE.frx":038A
      Left            =   8310
      List            =   "frmCollectionAE.frx":0394
      TabIndex        =   43
      Text            =   "On Hold"
      Top             =   120
      Width           =   2115
   End
   Begin Inventory.ctrlLiner ctrlLiner3 
      Height          =   30
      Left            =   120
      TabIndex        =   41
      Top             =   1380
      Width           =   10875
      _ExtentX        =   19182
      _ExtentY        =   53
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   5370
      TabIndex        =   24
      Top             =   150
      Width           =   1155
   End
   Begin VB.PictureBox ctrlLiner1 
      Height          =   30
      Left            =   -195
      ScaleHeight     =   30
      ScaleWidth      =   11340
      TabIndex        =   5
      Top             =   7485
      Width           =   11340
   End
   Begin VB.TextBox txtEntry 
      Height          =   990
      Index           =   3
      Left            =   90
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   8
      Tag             =   "Remarks"
      Top             =   6270
      Width           =   5805
   End
   Begin VB.TextBox txtTA 
      BackColor       =   &H00E6FFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9330
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0.00"
      Top             =   6060
      Width           =   1500
   End
   Begin VB.PictureBox picCusInfo 
      BorderStyle     =   0  'None
      Height          =   1650
      Left            =   75
      ScaleHeight     =   1650
      ScaleWidth      =   10965
      TabIndex        =   16
      Top             =   1560
      Width           =   10965
      Begin VB.ComboBox cbCA 
         Height          =   315
         ItemData        =   "frmCollectionAE.frx":03AC
         Left            =   1530
         List            =   "frmCollectionAE.frx":03B6
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   450
         Width           =   3315
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1530
         TabIndex        =   38
         Top             =   1230
         Width           =   1305
      End
      Begin VB.Frame Check 
         Caption         =   "Check"
         Height          =   1065
         Left            =   5010
         TabIndex        =   34
         Top             =   510
         Visible         =   0   'False
         Width           =   3825
         Begin VB.TextBox txtCheckNo 
            Height          =   315
            Left            =   1650
            TabIndex        =   35
            Top             =   210
            Width           =   1935
         End
         Begin ctrlNSDataCombo.NSDataCombo nsdBank 
            Height          =   315
            Left            =   1650
            TabIndex        =   50
            Top             =   600
            Width           =   1950
            _ExtentX        =   3440
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
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Bank:"
            Height          =   255
            Left            =   300
            TabIndex        =   37
            Top             =   630
            Width           =   1245
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Check Number:"
            Height          =   255
            Left            =   300
            TabIndex        =   36
            Top             =   270
            Width           =   1245
         End
      End
      Begin VB.ComboBox cbPT 
         Height          =   315
         ItemData        =   "frmCollectionAE.frx":03C8
         Left            =   1530
         List            =   "frmCollectionAE.frx":03D2
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   840
         Width           =   3315
      End
      Begin VB.CommandButton btnCollect 
         Caption         =   "Add To Collection"
         Enabled         =   0   'False
         Height          =   315
         Left            =   9210
         TabIndex        =   4
         Top             =   1230
         Width           =   1635
      End
      Begin VB.TextBox txtRem 
         ForeColor       =   &H00000000&
         Height          =   735
         Left            =   9000
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   390
         Width           =   1830
      End
      Begin VB.TextBox txtBal 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   3765
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   1275
         Width           =   1035
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdClient 
         Height          =   315
         Left            =   1530
         TabIndex        =   40
         Top             =   60
         Width           =   3300
         _ExtentX        =   5821
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
      Begin ctrlNSDataCombo.NSDataCombo nsdORNo 
         Height          =   315
         Left            =   6180
         TabIndex        =   46
         Top             =   60
         Width           =   2625
         _ExtentX        =   4630
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
         Caption         =   "Charge Account"
         Height          =   240
         Index           =   3
         Left            =   240
         TabIndex        =   49
         Top             =   450
         Width           =   1230
      End
      Begin VB.Label Labels 
         Caption         =   "DR#/ OR#:"
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
         Left            =   4980
         TabIndex        =   47
         Top             =   120
         Width           =   1050
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Amount:"
         Height          =   225
         Left            =   450
         TabIndex        =   39
         Top             =   1260
         Width           =   915
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Type"
         Height          =   240
         Index           =   12
         Left            =   330
         TabIndex        =   33
         Top             =   870
         Width           =   1140
      End
      Begin VB.Label Label5 
         Caption         =   "Remarks"
         ForeColor       =   &H0000011D&
         Height          =   240
         Left            =   9000
         TabIndex        =   22
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Balance"
         Height          =   240
         Index           =   6
         Left            =   3000
         TabIndex        =   21
         Top             =   1275
         Width           =   600
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer Name"
         Height          =   240
         Index           =   5
         Left            =   165
         TabIndex        =   17
         Top             =   60
         Width           =   1320
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   210
      Picture         =   "frmCollectionAE.frx":03E3
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Remove"
      Top             =   3840
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   8040
      TabIndex        =   11
      Top             =   7635
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9480
      TabIndex        =   12
      Top             =   7635
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   105
      TabIndex        =   10
      Top             =   7635
      Width           =   1755
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1425
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   2490
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   105
      TabIndex        =   6
      Top             =   3735
      Width           =   10755
      _ExtentX        =   18971
      _ExtentY        =   3863
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
   Begin VB.TextBox txtNCus 
      Height          =   210
      Left            =   6330
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   75
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1425
      TabIndex        =   23
      Top             =   525
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44630019
      CurrentDate     =   38207
   End
   Begin MSDataListLib.DataCombo dcBooking 
      Height          =   315
      Left            =   5370
      TabIndex        =   25
      Top             =   525
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcRoute 
      Height          =   315
      Left            =   1410
      TabIndex        =   26
      Top             =   900
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcCollection 
      Height          =   315
      Left            =   5370
      TabIndex        =   27
      Top             =   915
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin Inventory.ctrlLiner ctrlLiner2 
      Height          =   30
      Left            =   210
      TabIndex        =   42
      Top             =   3300
      Width           =   10785
      _ExtentX        =   19024
      _ExtentY        =   53
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   1425
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   525
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   285
      Left            =   7470
      TabIndex        =   44
      Top             =   150
      Width           =   795
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Collection"
      Height          =   255
      Left            =   4200
      TabIndex        =   31
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Booking"
      Height          =   255
      Left            =   4200
      TabIndex        =   30
      Top             =   525
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   240
      Left            =   150
      TabIndex        =   29
      Top             =   900
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Truck No"
      Height          =   255
      Left            =   4200
      TabIndex        =   28
      Top             =   135
      Width           =   1095
   End
   Begin VB.Label Labels 
      Caption         =   "Remarks"
      Height          =   240
      Index           =   4
      Left            =   120
      TabIndex        =   20
      Top             =   6000
      Width           =   990
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Total Collection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7230
      TabIndex        =   18
      Top             =   6060
      Width           =   2040
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Collection"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   180
      TabIndex        =   15
      Top             =   3435
      Width           =   4365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   14
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Collection No"
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
      Left            =   150
      TabIndex        =   13
      Top             =   150
      Width           =   1215
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   105
      Top             =   3435
      Width           =   10740
   End
End
Attribute VB_Name = "frmCollectionAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean

Dim cCAmount                As Currency 'Current Collection Amount
Dim cCRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice

Private Sub btnCollect_Click()
    If toNumber(txtPayment.Text) < 0 Then
        MsgBox "Please enter a valid payment.", vbExclamation
        txtPayment.SetFocus
        Exit Sub
    End If

'    Dim CurrRow As Integer
'
'    CurrRow = getFlexPos(Grid, 9, nsdClient.BoundText)

    'Add to grid
    With Grid
'        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 8) = "" Then
                .TextMatrix(1, 1) = nsdORNo.Text
                .TextMatrix(1, 2) = nsdClient.Text
                .TextMatrix(1, 3) = cbCA.Text
                .TextMatrix(1, 4) = cbPT.Text
                .TextMatrix(1, 5) = txtPayment.Text
                .TextMatrix(1, 6) = txtBal.Text
                .TextMatrix(1, 7) = txtRem.Text
                .TextMatrix(1, 8) = nsdClient.BoundText
                .TextMatrix(1, 9) = nsdORNo.BoundText
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdORNo.Text
                .TextMatrix(.Rows - 1, 2) = nsdClient.Text
                .TextMatrix(.Rows - 1, 3) = cbCA.Text
                .TextMatrix(.Rows - 1, 4) = cbPT.Text
                .TextMatrix(.Rows - 1, 5) = txtPayment.Text
                .TextMatrix(.Rows - 1, 6) = txtBal.Text
                .TextMatrix(.Rows - 1, 7) = txtRem.Text
                .TextMatrix(.Rows - 1, 8) = nsdClient.BoundText
                .TextMatrix(.Rows - 1, 9) = nsdORNo.BoundText

                .Row = .Rows - 1
            End If
            'Increase the record count
            cCRowCount = cCRowCount + 1
'        Else
'            If MsgBox("Payment already exist.Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
'                .Row = CurrRow
'
'                'Restore back the collected amount
'                cCAmount = cCAmount - toNumber(Grid.TextMatrix(.RowSel, 4))
'                txtTA.Text = toMoney(cCAmount)
'
'                'Replace collection
'                .TextMatrix(CurrRow, 1) = nsdORNo.Text
'                .TextMatrix(CurrRow, 2) = nsdClient.Text
'                .TextMatrix(CurrRow, 3) = cbPT.Text
'                .TextMatrix(CurrRow, 4) = txtPayment.Text
'                .TextMatrix(CurrRow, 5) = toMoney(txtBal) - toMoney(txtPayment.Text)
'                .TextMatrix(CurrRow, 6) = txtRem.Text
'                .TextMatrix(CurrRow, 7) = nsdClient.BoundText
'                .TextMatrix(CurrRow, 8) = nsdORNo.Text
'            Else
'                Exit Sub
'            End If
'        End If
        'Add the amount to current load amount
        cCAmount = cCAmount + toNumber(txtPayment.Text)
        txtTA.Text = toMoney(cCAmount)
        'Highlight the current row's column
        .ColSel = 8
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update amount to current collection amount
        cCAmount = cCAmount - toNumber(Grid.TextMatrix(.RowSel, 5))
        txtTA.Text = toMoney(cCAmount)
        'Update the record count
        cCRowCount = cCRowCount - 1

        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cbPT_Click()
    If cbPT.ListIndex = 0 Then
        Check.Visible = False
    Else
        Check.Visible = True
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdExpenses_Click()
    With frmSalesExpenses
        .PK = PK
        
        .show 1
    End With
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
       
    If dcBooking.BoundText = "" Then
        MsgBox "Please select a booking agent.", vbExclamation
        dcBooking.SetFocus
        Exit Sub
    End If
    
    If dcCollection.BoundText = "" Then
        MsgBox "Please select a collection agent.", vbExclamation
        dcCollection.SetFocus
        Exit Sub
    End If

    If cCRowCount < 1 Then
        MsgBox "Please enter a collection first before saving this record.", vbExclamation
        txtEntry(0).SetFocus
        Exit Sub
    End If

    If MsgBox("This save the record.Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    
    rs.Close
    rs.Open "SELECT * FROM Receipts_Batch WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Collection_Details WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    DeleteItems
    
    'Save the record
    With rs

        ![Status] = IIf(cboStatus.Text = "Collected", True, False)
        ![Remarks] = txtEntry(3).Text
        ![Gross] = toMoney(txtTA.Text)
        
        .Update
    End With

    With Grid
        'Save the details of the records
        For c = 1 To cCRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                RSDetails.AddNew

                RSDetails![ReceiptBatchID] = PK
                RSDetails![RefNo] = .TextMatrix(c, 1)
                RSDetails![ClientID] = .TextMatrix(c, 8)
                RSDetails![ChargeAccount] = .TextMatrix(c, 3)
                RSDetails![PaymentType] = .TextMatrix(c, 4)
                RSDetails![Amount] = toNumber(.TextMatrix(c, 5))
                RSDetails![Balance] = toNumber(.TextMatrix(c, 6))
                RSDetails![Remarks] = .TextMatrix(c, 7)
                RSDetails![ReceiptID] = .TextMatrix(c, 9)
         
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "CollectionDetailID = " & toNumber(.TextMatrix(c, 10))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew
                
                RSDetails![ReceiptBatchID] = PK
                RSDetails![RefNo] = .TextMatrix(c, 1)
                RSDetails![ClientID] = .TextMatrix(c, 8)
                RSDetails![ChargeAccount] = .TextMatrix(c, 3)
                RSDetails![PaymentType] = .TextMatrix(c, 4)
                RSDetails![Amount] = toNumber(.TextMatrix(c, 5))
                RSDetails![Balance] = toNumber(.TextMatrix(c, 6))
                RSDetails![Remarks] = .TextMatrix(c, 7)
                RSDetails![ReceiptID] = .TextMatrix(c, 9)
         
                RSDetails.Update
            End If
            
            If cboStatus.Text = "Collected" Then
                Dim LedgerID As Integer
                
                LedgerID = getIndex("Clients_Ledger")
                CN.Execute "INSERT INTO Clients_Ledger (LedgerID, ReceiptID, ReceiptBatchID, ClientID, [Date], RefNo, ChargeAccount, PaymentType, Debit, Credit ) " _
                        & "VALUES (" & LedgerID & ", " & .TextMatrix(c, 9) & ", " & PK & ", " & .TextMatrix(c, 8) & ", #" & dtpDate.Value & "#, '" & .TextMatrix(c, 1) & "', '" & .TextMatrix(c, 3) & "' , '" & .TextMatrix(c, 4) & "', 0, " & toNumber(.TextMatrix(c, 5)) & ")"
                        
                If cbPT.ListIndex = 1 Then
                    CN.Execute "INSERT INTO Payments_Checks ( LedgerID, CheckNo, BankID, [Date] ) " _
                            & "VALUES (" & LedgerID & ", " & txtCheckNo.Text & ", " & nsdBank.BoundText & ", Date())"
                End If
            End If
            
        Next c
    End With

    'Clear variables
    c = 0
    
    Set RSDetails = Nothing

    CN.CommitTrans

    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me

    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then Unload Me: Exit Sub
    txtEntry(0).SetFocus
End Sub

Private Sub Form_Load()
    
    'Bind the data combo
    bind_dc "SELECT * FROM Routes", "Desc", dcRoute, "RouteID", True
    bind_dc "SELECT * FROM Agents", "AgentCode", dcBooking, "AgentID", True
    bind_dc "SELECT * FROM Agents", "AgentCode", dcCollection, "AgentID", True

    InitGrid
    InitNSD

    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Initialize controls
        cbPT.ListIndex = 0

        'Set the recordset
         rs.Open "SELECT * FROM Receipts_Batch WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         Caption = "Create New Entry"
         cmdUsrHistory.Enabled = False
         GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_Receipts_Batch WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If State = adStateViewMode Then
            Caption = "Edit Record"
            cmdCancel.Caption = "Close"
            DisplayForViewing
        Else
            Caption = "Edit Record"
            cmdCancel.Caption = "Cancel"
            DisplayForEditing
        End If
        
        cmdUsrHistory.Enabled = True
        
        Screen.MousePointer = vbDefault
    End If
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Collection")
    txtEntry(0).Text = "COL" & GenerateID(PK, Format$(Date, "yyyy") & Format$(Date, "mm") & Format$(Date, "dd") & "-", "0")
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cCRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 11
        .ColSel = 7
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1000
        .ColWidth(2) = 3500
        .ColWidth(3) = 1500
        .ColWidth(4) = 1500
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        .ColWidth(7) = 2000
        .ColWidth(8) = 0
        .ColWidth(9) = 0
        .ColWidth(10) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "OR No"
        .TextMatrix(0, 2) = "Customer Name"
        .TextMatrix(0, 3) = "Charge Account"
        .TextMatrix(0, 4) = "Payment Type"
        .TextMatrix(0, 5) = "Payment"
        .TextMatrix(0, 6) = "Balance"
        .TextMatrix(0, 7) = "Remarks"
        .TextMatrix(0, 8) = "ClientID"
        .TextMatrix(0, 9) = "ReceiptID"
        .TextMatrix(0, 10) = "CollectionDetailID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbLeftJustify
'        .ColAlignment(5) = vbLeftJustify
'        .ColAlignment(6) = vbLeftJustify
    End With
End Sub

Private Sub ResetEntry()
    nsdClient.ResetValue
    txtBal.Text = "0.00"
    
    txtPayment.Text = "0.00"
    cbPT.ListIndex = 0
    txtRem.Text = ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmCollection.RefreshRecords
    End If
    
    Set frmCollectionAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateViewMode Then Exit Sub
    
    If Grid.Rows = 2 And Grid.TextMatrix(1, 8) = "" Then '8 = ClientID
        btnRemove.Visible = False
    Else
        btnRemove.Visible = True
        btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
        btnRemove.Left = Grid.Left + 50
    End If
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub nsdClient_Change()
    nsdORNo.sqlwCondition = "ClientID=" & nsdClient.BoundText
    
    txtBal.Text = nsdClient.getSelValueAt(3)
    txtBal.Tag = nsdClient.getSelValueAt(3)
    
    txtPayment.Text = "0.00"
End Sub

Private Sub nsdORNo_Change()
    txtPayment.Text = nsdORNo.getSelValueAt(3)
    
    Dim dDeliveryDate As Date
    
    dDeliveryDate = getValueAt("SELECT LedgerID, Date FROM Clients_Ledger WHERE RefNo = '" & nsdORNo.Text & "'", "Date")
    
    If dtpDate.Value = dDeliveryDate Then
        cbCA.ListIndex = 0
    Else
        cbCA.ListIndex = 1
    End If
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtPayment_Change()
    If toNumber(txtPayment.Text) > 0 Then
        btnCollect.Enabled = True
    Else
        btnCollect.Enabled = False
    End If
    
    If toNumber(txtPayment.Text) > toNumber(txtBal.Tag) Then
        txtBal.Text = "0.00"
        txtPayment.Text = toMoney(toNumber(txtBal.Tag))
        txtPayment.SelStart = Len(txtPayment.Text)
    Else
        txtBal.Text = toMoney(toNumber(txtBal.Tag) - toNumber(txtPayment.Text))
    End If
End Sub

Private Sub txtPayment_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtPayment_Validate(Cancel As Boolean)
    txtPayment.Text = toMoney(toNumber(txtPayment.Text))
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
    If Index = 3 Then
        cmdSave.Default = False
    End If
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index > 1 And Index < 3 Then
        KeyAscii = isNumber(KeyAscii)
    End If
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 3 Then
        cmdSave.Default = True
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index > 1 And Index < 3 Then
        txtEntry(Index).Text = toNumber(txtEntry(Index).Text)
    End If
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry

    dtpDate.Value = Date

    txtEntry(3).Text = ""

    txtTA.Text = "0.00"

    cCAmount = 0

    txtEntry(0).SetFocus
End Sub

'Used to display record
Private Sub DisplayForEditing()
    On Error GoTo err
    txtEntry(0).Text = rs![ReceiptBatchID]
    dtpDate.Value = Format$(rs![DeliveryDate], "MMM-dd-yyyy")
    dcRoute.BoundText = rs![RouteID]
    txtEntry(1).Text = rs![TruckNo]
    dcBooking.BoundText = rs![BookingAgent]
    dcCollection.BoundText = rs![CollectionAgent]
    txtEntry(3).Text = rs![Remarks]
    txtTA.Text = toMoney(toNumber(rs![Gross]))
    cboStatus.Text = rs!Status_Alias
    
    'Display the details
    Dim RSDetails As New Recordset

    cCAmount = txtTA.Text
    cCRowCount = 0
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Collection_Details WHERE ReceiptBatchID=" & PK & " ORDER BY CollectionDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cCRowCount = cCRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 9) = "" Then
                    .TextMatrix(1, 1) = RSDetails![RefNo]
                    .TextMatrix(1, 2) = RSDetails![Company]
                    .TextMatrix(1, 3) = RSDetails![ChargeAccount]
                    .TextMatrix(1, 4) = RSDetails![PaymentType]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(1, 7) = RSDetails![Remarks]
                    .TextMatrix(1, 8) = RSDetails![ClientID]
                    .TextMatrix(1, 9) = RSDetails![ReceiptID]
                    .TextMatrix(1, 10) = RSDetails![CollectionDetailID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![RefNo]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Company]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![ChargeAccount]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![PaymentType]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Remarks]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![ClientID]
                    .TextMatrix(.Rows - 1, 9) = RSDetails![ReceiptID]
                    .TextMatrix(.Rows - 1, 10) = RSDetails![CollectionDetailID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 8
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then Resume Next
    
    MsgBox err.Number & " " & err.Description
End Sub

Private Sub DisplayForViewing()
    On Error GoTo err
    txtEntry(0).Text = rs![ReceiptBatchID]
    dtpDate.Value = Format$(rs![DeliveryDate], "MMM-dd-yyyy")
    dcRoute.BoundText = rs![RouteID]
    txtEntry(1).Text = rs![TruckNo]
    dcBooking.BoundText = rs![Booking]
    dcCollection.BoundText = rs![Collection]
    txtEntry(3).Text = rs![Remarks]
    txtTA.Text = toMoney(toNumber(rs![Gross]))
    cboStatus.Text = rs!Status_Alias
    
    'Display the details
    Dim RSDetails As New Recordset

    cCAmount = txtTA.Text
    cCRowCount = 0
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Collection_Details WHERE ReceiptBatchID=" & PK & " ORDER BY CollectionDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cCRowCount = cCRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 8) = "" Then
                    .TextMatrix(1, 1) = RSDetails![RefNo]
                    .TextMatrix(1, 2) = RSDetails![Company]
                    .TextMatrix(1, 3) = RSDetails![ChargeAccount]
                    .TextMatrix(1, 4) = RSDetails![PaymentType]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(1, 7) = RSDetails![Remarks]
                    .TextMatrix(1, 8) = RSDetails![ClientID]
'                    .TextMatrix(1, 8) = RSDetails![ReceiptID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![RefNo]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Company]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![ChargeAccount]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![PaymentType]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Amount])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Balance])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Remarks]
'                    .TextMatrix(.Rows - 1, 7) = RSDetails![ClientID]
'                    .TextMatrix(.Rows - 1, 8) = RSDetails![ReceiptID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 8
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    'Disable commands
    LockInput Me, True

    picCusInfo.Visible = False
    dtpDate.Visible = False
    txtDate.Visible = True
    cmdSave.Visible = False
    btnCollect.Visible = False
    btnRemove.Visible = False
    
    'Resize and reposition the controls
    Shape3.Top = 1350
    Label11.Top = 1350
    Grid.Top = 1600
    Grid.Height = 3590

    ctrlLiner2.Visible = False
    ctrlLiner3.Visible = False

    Label3.Top = 5425
    txtTA.Top = 5425

    Labels(4).Top = 5325
    txtEntry(3).Top = 5550

    ctrlLiner1.Top = 6800

    cmdUsrHistory.Top = 6950
    cmdCancel.Top = 6950

    Me.Height = 7900
    Me.Top = (Screen.Height - Me.Height) / 2

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then Resume Next
    
    MsgBox err.Number & " " & err.Description
End Sub

Private Sub txtTA_GotFocus()
    HLText txtTA
End Sub

Private Sub txtPayment_GotFocus()
    HLText txtPayment
End Sub

Private Sub InitNSD()
    'For Invoice
    With nsdClient
        .ClearColumn
        .AddColumn "Company", 1794.89
        .AddColumn "OwnersName", 1994.89
        .AddColumn "Balance", 2264.88
        
        .Connection = CN.ConnectionString
        .sqlFields = "Company,OwnersName,Balance,ClientID"
        .sqlTables = "qry_Client_Balance"
        .sqlSortOrder = "Company ASC"
        
        .BoundField = "ClientID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Unpaid Invoices"
    End With

    'For DR#/OR#
    With nsdORNo
        .ClearColumn
        .AddColumn "OR No", 1300.89
        .AddColumn "Company", 1994.89
        .AddColumn "Balance", 1264.88
        .AddColumn "ReceiptID", 0
        
        .Connection = CN.ConnectionString
        .sqlFields = "RefNo,Company,Balance,ReceiptID"
        .sqlTables = "qry_Client_Balance_Per_Receipt"
        
        .sqlSortOrder = "RefNo ASC"
        
        .BoundField = "ReceiptID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6500, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Unpaid Invoices"
    End With
    
    'For Bank
    With nsdBank
        .ClearColumn
        .AddColumn "Bank ID", 1794.89
        .AddColumn "Bank", 2264.88
        .AddColumn "Branch", 2670.23
        .Connection = CN.ConnectionString
        
        '.sqlFields = "VendorID, Company, Location"
        .sqlFields = "BankID, Bank, Branch"
        .sqlTables = "Banks"
        .sqlSortOrder = "Bank ASC"
        
        .BoundField = "BankID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Banks Record"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSCollection As New Recordset

    If State = adStateAddMode Then Exit Sub

    RSCollection.CursorLocation = adUseClient
    RSCollection.Open "SELECT * FROM Collection_Details WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSCollection.RecordCount > 0 Then
        RSCollection.MoveFirst
        While Not RSCollection.EOF
            CurrRow = getFlexPos(Grid, 8, RSCollection!ClientID)

            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Collection_Details", "CollectionDetailID", "", True, RSCollection!CollectionDetailID
                End If
            End With
            RSCollection.MoveNext
        Wend
    End If
End Sub
