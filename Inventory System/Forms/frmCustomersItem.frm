VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmCustomersItem 
   BorderStyle     =   0  'None
   Caption         =   "Customer's Items"
   ClientHeight    =   8565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12915
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   375
      Left            =   11550
      TabIndex        =   32
      Top             =   6840
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   11550
      TabIndex        =   31
      Top             =   7290
      Width           =   1125
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9810
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7080
      Width           =   1425
   End
   Begin VB.TextBox txtGross 
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
      Index           =   2
      Left            =   9810
      Locked          =   -1  'True
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6780
      Width           =   1425
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   9810
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   8130
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9810
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7380
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   9810
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7680
      Width           =   1425
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   150
      ScaleHeight     =   630
      ScaleWidth      =   12525
      TabIndex        =   0
      Top             =   3150
      Width           =   12525
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   255
         Width           =   1290
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   4470
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   2775
         TabIndex        =   5
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10380
         TabIndex        =   4
         Top             =   255
         Width           =   840
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   9420
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   270
         Width           =   855
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   8580
         TabIndex        =   2
         Text            =   "0"
         Top             =   270
         Width           =   735
      End
      Begin VB.TextBox txtExtPrice 
         Height          =   285
         Left            =   5700
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   240
         Width           =   1185
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   225
         Width           =   2640
         _ExtentX        =   4657
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
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   9
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   7155
         TabIndex        =   17
         Top             =   30
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   16
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   15
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Items/Stocks"
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
         Index           =   8
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   13
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9420
         TabIndex        =   12
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   8580
         TabIndex        =   11
         Top             =   30
         Width           =   690
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Price"
         Height          =   240
         Index           =   5
         Left            =   5730
         TabIndex        =   10
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   150
      TabIndex        =   18
      Top             =   4080
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   4604
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
   Begin MSComctlLib.ListView lvList 
      Height          =   2760
      Left            =   150
      TabIndex        =   20
      Top             =   120
      Width           =   12525
      _ExtentX        =   22093
      _ExtentY        =   4868
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
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
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   6068
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Qty"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ImageList i16x16 
      Left            =   3210
      Top             =   7110
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":0A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":1424
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":1E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":2848
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCustomersItem.frx":325A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7710
      TabIndex        =   30
      Top             =   7110
      Width           =   2040
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Gross"
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
      Left            =   7710
      TabIndex        =   29
      Top             =   6810
      Width           =   2040
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7710
      TabIndex        =   28
      Top             =   8160
      Width           =   2040
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Tax Base"
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
      Left            =   7710
      TabIndex        =   27
      Top             =   7410
      Width           =   2040
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Vat(0.12)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000011D&
      Height          =   240
      Left            =   7710
      TabIndex        =   26
      Top             =   7710
      Width           =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   9480
      X2              =   11190
      Y1              =   8070
      Y2              =   8070
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipt Details"
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
      Left            =   240
      TabIndex        =   19
      Top             =   3840
      Width           =   4365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   150
      X2              =   12650
      Y1              =   3030
      Y2              =   3030
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   150
      X2              =   12700
      Y1              =   3060
      Y2              =   3060
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   3810
      Width           =   12525
   End
End
Attribute VB_Name = "frmCustomersItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public StockID              As Integer
Public blnCancel            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim rs                      As New Recordset
Dim intQtyOld               As Integer

Private Sub btnUpdate_Click()
On Error GoTo err
   
    If MsgBox("Are you sure you want to update this record?" & vbCrLf & vbCrLf & _
            "Click Yes to proceed", vbInformation + vbYesNo) = vbNo Then
        txtQty.Text = intQtyOld
        Exit Sub
    End If
    
'    CN.BeginTrans
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
       
    CurrRow = getFlexPos(Grid, 10, nsdStock.Tag)
    intStockID = nsdStock.Tag
   
    ChangeValue CN, "Receipts_Detail", "Qty", txtQty.Text, False, "ReceiptID=" & lvList.SelectedItem.Tag & " AND StockID=" & intStockID & ""
    ChangeValue CN, "Receipts", "Deducted", "Yes", False, "ReceiptID=" & lvList.SelectedItem.Tag & ""
    
    Dim RSStockUnit As New Recordset

    With RSStockUnit
        .CursorLocation = adUseClient
        .Open "SELECT * FROM Stock_Unit WHERE StockID =" & intStockID & " AND UnitID = " & toNumber(dcUnit.BoundText), CN, adOpenStatic, adLockOptimistic
        
        !Onhand = !Onhand + (intQtyOld - toNumber(txtQty.Text))
        
        .Update
    End With
    
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT ReceiptDetailID, Qty FROM Receipts_Detail WHERE ReceiptID=" & lvList.SelectedItem.Tag & " AND StockID =" & intStockID & " ORDER BY ReceiptDetailID ASC", CN, adOpenStatic, adLockOptimistic
    
    RSDetails![Qty] = toNumber(txtQty.Text)
    
    RSDetails.Update
    
    'Add to grid
    With Grid
        .Row = CurrRow
        
        'Restore back the invoice amount and discount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 6))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 8))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.RowSel, 3)) * toNumber(txtPrice.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        
        .TextMatrix(CurrRow, 1) = nsdStock.getSelValueAt(1)
        .TextMatrix(CurrRow, 2) = nsdStock.Text
        .TextMatrix(CurrRow, 3) = toNumber(txtQty.Text) 'txtQty.Text
        .TextMatrix(CurrRow, 4) = dcUnit.Text
        .TextMatrix(CurrRow, 5) = toMoney(txtPrice.Text)
        .TextMatrix(CurrRow, 6) = toMoney(txtExtPrice.Text)
        .TextMatrix(CurrRow, 7) = toMoney(txtGross(1).Text)
        .TextMatrix(CurrRow, 8) = toNumber(txtDisc.Text)
        .TextMatrix(CurrRow, 9) = toMoney(toNumber(txtNetAmount.Text))
        .TextMatrix(CurrRow, 10) = intStockID

        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        cDAmount = cDAmount + toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(txtQty.Text) * toNumber(txtPrice.Text)))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        'Highlight the current row's column
        .ColSel = 10
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
        
'    CN.CommitTrans
    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdCancel_Click()
    blnCancel = True
    
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Receipts_Detail WHERE ReceiptID=" & lvList.SelectedItem.Tag & " ORDER BY ReceiptDetailID ASC, CN, adOpenStatic, adLockOptimistic"
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c

            RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 10))
        
            If RSDetails.RecordCount = 0 Then MsgBox "Error! Current record has been deleted.", vbExclamation: Exit Sub
            
            RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
            
            RSDetails.Update
        Next c
    End With
    
    If cIRowCount > 0 Then _
        ChangeValue CN, "Receipts", "Deducted", "True", False, "ReceiptID=" & lvList.SelectedItem.Tag & ""
        
    Set RSDetails = Nothing
    
    blnCancel = False
    Unload Me
    
    Exit Sub

err:
    MsgBox err.Description, vbInformation
    blnCancel = False
End Sub

Private Sub cmdOK_Click()
    blnCancel = False
    Unload Me
End Sub

Private Sub Form_Load()
    If rs.State = adStateOpen Then rs.Close
    
    InitGrid
    
    rs.Open "SELECT Company,TotalQty,Gross,Discount,TaxBase,Vat,NetAmount,ReceiptID FROM qry_Receipts_Qty WHERE StockID = " & StockID & " ORDER BY ReceiptID ASC", CN, adOpenStatic, adLockOptimistic
    FillListView lvList, rs, 2, 0, False, True, "ReceiptID"
End Sub

Private Sub Grid_Click()
    With Grid
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 10), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        nsdStock.Text = .TextMatrix(.RowSel, 2)
        nsdStock.Tag = .TextMatrix(.RowSel, 10) 'Add tag coz boundtext is empty
        intQtyOld = .TextMatrix(.RowSel, 3)
        txtQty.Text = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtPrice.Text = toMoney(.TextMatrix(.RowSel, 5))
        txtExtPrice.Text = toMoney(.TextMatrix(.RowSel, 6))
        txtGross(1).Text = toMoney(.TextMatrix(.RowSel, 7))
        txtDisc.Text = toMoney(.TextMatrix(.RowSel, 8))
        txtNetAmount.Text = toMoney(.TextMatrix(.RowSel, 9))
    End With
End Sub

Private Sub lvList_Click()
    rs.MoveFirst
    rs.Find "ReceiptID = " & lvList.SelectedItem.Tag
    
    DisplayForEditing
End Sub

'Used to display record
Private Sub DisplayForEditing()
    On Error GoTo err
       
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])

    cIGross = toNumber(txtGross(2).Text)
    cDAmount = toNumber(txtDesc.Text)
    cIAmount = toNumber(txtNet.Text)
'    cIRowCount = 1
    
    Dim I As Integer
    
    For I = 1 To cIRowCount
        If Grid.Rows = 2 Then Grid.Rows = Grid.Rows + 1
        Grid.RemoveItem (Grid.RowSel)
    Next
    
    cIRowCount = 0

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Receipts_Detail WHERE ReceiptID=" & lvList.SelectedItem.Tag & " ORDER BY ReceiptDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![ExtPrice])
                    .TextMatrix(1, 7) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 10) = RSDetails![StockID]
                    .TextMatrix(1, 11) = RSDetails![Suggested]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![ExtPrice]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                    .TextMatrix(.Rows - 1, 11) = RSDetails![Suggested]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 10
        'Set fixed cols
'        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
'        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
    
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 12
        .ColSel = 11
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2025
        .ColWidth(2) = 2505
        .ColWidth(3) = 1545
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 900
        .ColWidth(8) = 900
        .ColWidth(9) = 900
        .ColWidth(10) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Sales Price"
        .TextMatrix(0, 6) = "Ext Price"
        .TextMatrix(0, 7) = "Gross"
        .TextMatrix(0, 8) = "Discount(%)"
        .TextMatrix(0, 9) = "Net Amount"
        .TextMatrix(0, 10) = "Stock ID"
        .TextMatrix(0, 11) = "Suggested"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbRightJustify
        .ColAlignment(5) = vbLeftJustify
        .ColAlignment(6) = vbRightJustify
        .ColAlignment(7) = vbRightJustify
        .ColAlignment(8) = vbRightJustify
        .ColAlignment(9) = vbRightJustify
    End With
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnUpdate.Enabled = False
        Exit Sub
    Else
        btnUpdate.Enabled = True
    End If
       
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtPrice.Tag = 0
    txtPrice.Text = "0.00"
    txtQty.Text = 0
End Sub
