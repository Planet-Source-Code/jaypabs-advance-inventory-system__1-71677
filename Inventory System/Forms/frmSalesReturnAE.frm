VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmSalesReturnAE 
   BorderStyle     =   0  'None
   ClientHeight    =   9030
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12345
   LinkTopic       =   "Form1"
   ScaleHeight     =   9030
   ScaleWidth      =   12345
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7950
      Width           =   1425
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   25
      Top             =   8490
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   10755
      TabIndex        =   24
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   9300
      TabIndex        =   23
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmSalesReturnAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Remove"
      Top             =   4200
      Visible         =   0   'False
      Width           =   275
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
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   225
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Tag             =   "Remarks"
      Top             =   6870
      Width           =   5910
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10680
      Locked          =   -1  'True
      TabIndex        =   19
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6900
      Width           =   1425
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   210
      ScaleHeight     =   630
      ScaleWidth      =   11910
      TabIndex        =   3
      Top             =   3360
      Width           =   11910
      Begin VB.ComboBox cboReturnType 
         Height          =   315
         ItemData        =   "frmSalesReturnAE.frx":01B2
         Left            =   9810
         List            =   "frmSalesReturnAE.frx":01BC
         TabIndex        =   44
         Text            =   "Bad"
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtDisc 
         Height          =   285
         Left            =   8160
         Locked          =   -1  'True
         TabIndex        =   10
         Text            =   "0"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   8910
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10980
         TabIndex        =   8
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   285
         Left            =   4215
         TabIndex        =   7
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.TextBox txtPrice 
         Height          =   285
         Left            =   5970
         Locked          =   -1  'True
         TabIndex        =   6
         Text            =   "0.00"
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   7035
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   240
         Width           =   1080
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   4950
         TabIndex        =   4
         Top             =   240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   60
         TabIndex        =   11
         Top             =   210
         Width           =   4080
         _ExtentX        =   7197
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
      Begin VB.Label Label5 
         Caption         =   "Type of Return"
         Height          =   195
         Left            =   9840
         TabIndex        =   45
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   8100
         TabIndex        =   18
         Top             =   0
         Width           =   765
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   8910
         TabIndex        =   17
         Top             =   0
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   4950
         TabIndex        =   16
         Top             =   0
         Width           =   900
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
         Left            =   60
         TabIndex        =   15
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   5940
         TabIndex        =   14
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   4215
         TabIndex        =   13
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   7035
         TabIndex        =   12
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.TextBox txtCity 
      Height          =   315
      Left            =   1710
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   3315
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmSalesReturnAE.frx":01CB
      Left            =   7530
      List            =   "frmSalesReturnAE.frx":01D5
      TabIndex        =   1
      Text            =   "On Hold"
      Top             =   900
      Width           =   2325
   End
   Begin VB.TextBox txtReturnSlipNo 
      Height          =   315
      Left            =   1710
      TabIndex        =   0
      Top             =   1980
      Width           =   2625
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   27
      Top             =   8340
      Width           =   11820
      _ExtentX        =   20849
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2490
      Left            =   180
      TabIndex        =   28
      Top             =   4050
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   4392
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   1710
      TabIndex        =   29
      Top             =   1590
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "dd MMM, yyyy"
      Format          =   44630019
      CurrentDate     =   38989
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdClient 
      Height          =   315
      Left            =   1710
      TabIndex        =   43
      Top             =   840
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
   Begin VB.TextBox txtDate 
      Height          =   314
      Left            =   1710
      TabIndex        =   30
      Top             =   1590
      Width           =   1935
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   390
      TabIndex        =   42
      Top             =   840
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5100
      TabIndex        =   41
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   10350
      X2              =   12060
      Y1              =   7890
      Y2              =   7890
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   3150
      Width           =   11910
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
      Left            =   8580
      TabIndex        =   40
      Top             =   7980
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   12090
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   12060
      Y1              =   3060
      Y2              =   3060
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
      Left            =   8580
      TabIndex        =   39
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   38
      Top             =   6600
      Width           =   990
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8895
      Left            =   60
      Top             =   30
      Width           =   12225
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
      Left            =   8580
      TabIndex        =   37
      Top             =   6930
      Width           =   2040
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Return"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   240
      TabIndex        =   36
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order Return Details"
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
      Left            =   210
      TabIndex        =   35
      Top             =   3150
      Width           =   4365
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   255
      Left            =   930
      TabIndex        =   34
      Top             =   1590
      Width           =   705
   End
   Begin VB.Shape Shape1 
      Height          =   8235
      Left            =   150
      Top             =   600
      Width           =   12045
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "City"
      Height          =   285
      Left            =   390
      TabIndex        =   33
      Top             =   1200
      Width           =   1245
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   6180
      TabIndex        =   32
      Top             =   930
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Return Slip No"
      Height          =   255
      Left            =   390
      TabIndex        =   31
      Top             =   2010
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   12075
   End
End
Attribute VB_Name = "frmSalesReturnAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public ReceiptPK            As Long 'Variable used to save the reference from SeceiptID
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim intQtyOld               As Integer 'Allowed value for return qty
Dim blnSave                 As Boolean

Private Sub btnAdd_Click()
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency

    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub

    If dcUnit.Text = "" Then
        MsgBox "Please select unit", vbInformation
        dcUnit.SetFocus
        Exit Sub
    End If

    If toNumber(txtPrice.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtPrice.SetFocus
        Exit Sub
    End If

    Dim CurrRow As Integer
    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 10, nsdStock.Tag)
    intStockID = nsdStock.Tag

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                .TextMatrix(1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(1, 2) = nsdStock.Text
                .TextMatrix(1, 3) = txtQty.Text
                .TextMatrix(1, 4) = dcUnit.Text
                .TextMatrix(1, 5) = toMoney(txtPrice.Text)
                .TextMatrix(1, 6) = toMoney(txtGross(1).Text)
                .TextMatrix(1, 7) = toMoney(txtDisc.Text)
                .TextMatrix(1, 8) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(1, 9) = cboReturnType.Text
                .TextMatrix(1, 10) = intStockID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(.Rows - 1, 2) = nsdStock.Text
                .TextMatrix(.Rows - 1, 3) = txtQty.Text
                .TextMatrix(.Rows - 1, 4) = dcUnit.Text
                .TextMatrix(.Rows - 1, 5) = toMoney(txtPrice.Text)
                .TextMatrix(.Rows - 1, 6) = toMoney(txtGross(1).Text)
                .TextMatrix(.Rows - 1, 7) = toMoney(txtDisc.Text)
                .TextMatrix(.Rows - 1, 8) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(.Rows - 1, 9) = cboReturnType.Text
                .TextMatrix(.Rows - 1, 10) = intStockID
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
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
                .TextMatrix(CurrRow, 3) = txtQty.Text
                .TextMatrix(CurrRow, 4) = dcUnit.Text
                .TextMatrix(CurrRow, 5) = toMoney(txtPrice.Text)
                .TextMatrix(CurrRow, 6) = toMoney(txtGross(1).Text)
                .TextMatrix(CurrRow, 7) = toMoney(txtDisc.Text)
                .TextMatrix(CurrRow, 8) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(CurrRow, 9) = cboReturnType.Text
                .TextMatrix(CurrRow, 10) = intStockID

            Else
                Exit Sub
            End If
        End If
        
        'Save to stock card
        Dim RSStockCard As New Recordset
    
        RSStockCard.CursorLocation = adUseClient
        RSStockCard.Open "Stock_Card", CN, , adLockOptimistic, adCmdTable
        
        'Add record to stock card
        RSStockCard.AddNew
            
        RSStockCard!Type = "POR"
        RSStockCard!RefNo1 = ReceiptPK
        RSStockCard!Pieces1 = "-" & toNumber(txtQty.Text)
        RSStockCard!Cost = toNumber(txtPrice.Text)
        RSStockCard!StockID = intStockID
            
        RSStockCard.Update
            
        'Deduct qty returned to qty onhand in Stock_Unit tables
        Dim RSStockUnit As New Recordset
    
        RSStockUnit.CursorLocation = adUseClient
        RSStockUnit.Open "SELECT * From Stock_Unit", CN, adOpenStatic, adLockOptimistic
            
        'Deduct qty returned in stocks table
        RSStockUnit.Filter = "StockID = " & intStockID & " AND UnitID = " & dcUnit.BoundText
        
        RSStockUnit!Onhand = RSStockUnit!Onhand - toNumber(txtQty.Text)
        
        RSStockUnit.Update
                    
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        
        cDAmount = cDAmount + toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(txtQty.Text) * toNumber(txtPrice.Text)))
        
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        
        'Highlight the current row's column
        .ColSel = 9
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update grooss to current purchase amount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 6))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 8))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.RowSel, 4)) * toNumber(Grid.TextMatrix(.RowSel, 6))))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")

        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
    
End Sub

Private Sub dcUnit_Change()
    If dcUnit.Text = "" Or nsdStock.Tag = "" Then Exit Sub
    
    txtPrice.Text = toMoney(getValueAt("SELECT SalesPrice FROM qry_Stock_Unit WHERE StockID= " & nsdStock.Tag & " AND UnitID = " & dcUnit.BoundText & "", "SalesPrice"))
End Sub

Private Sub nsdClient_Change()
    txtCity.Text = nsdClient.getSelValueAt(3)
End Sub

Private Sub nsdStock_Change()
    On Error Resume Next
    
    nsdStock.Tag = nsdStock.BoundText
    txtQty.Text = "0"
    
    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText & " ORDER BY qry_Unit.Order ASC", "Unit", dcUnit, "UnitID", True
    
'    txtPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Supplier Price
End Sub

Private Sub txtdisc_Change()
    txtQty_Change
End Sub

Private Sub txtdisc_Click()
    txtQty_Change
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    
    Unload Me
End Sub

Private Sub txtDisc_GotFocus()
    HLText txtDisc
End Sub

Private Sub txtdisc_Validate(Cancel As Boolean)
    txtDisc.Text = toNumber(txtDisc.Text)
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If txtReturnSlipNo.Text = "" Then
        MsgBox "Please don't leave Return Slip No field blank.", vbInformation
        txtReturnSlipNo.SetFocus
        Exit Sub
    End If
    
    If cIRowCount < 1 Then
        MsgBox "Please enter item to return before saving this record.", vbExclamation
        Exit Sub
    End If
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Local_Purchase_Return
    Dim RSReturn As New Recordset

    RSReturn.CursorLocation = adUseClient
    RSReturn.Open "Sales_Return", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    'Connection for Purchase_Order_Return_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "Sales_Return_Detail WHERE SalesReturnID=" & PK, CN, adOpenDynamic, adLockOptimistic, adCmdTable

    Screen.MousePointer = vbHourglass

    Dim c As Integer
    
    DeleteItems
    
    On Error GoTo err

    'Save the record
    With RSReturn
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![SalesReturnID] = PK
            ![ClientID] = nsdClient.Tag
            ![ReceiptID] = ReceiptPK
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        ElseIf State = adStateEditMode Then
            .Close
            .Open "SELECT * FROM Sales_Return WHERE SalesReturnID=" & PK, CN, adOpenStatic, adLockOptimistic
            
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If

        ![ReturnSlipNo] = txtReturnSlipNo.Text
        ![Date] = dtpDate.Value
        ![Status] = IIf(cboStatus.Text = "Returned", True, False)
        ![Notes] = txtNotes.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![NetAmount] = toNumber(txtNet.Text)
                
        .Update
    End With
   
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSDetails.AddNew
            
                RSDetails![SalesReturnID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 10))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 4))
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ReturnType] = .TextMatrix(c, 9)
    
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 10))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew
            
                RSDetails![StockID] = toNumber(.TextMatrix(c, 10))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 4))
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ReturnType] = .TextMatrix(c, 9)
    
                RSDetails.Update
            End If
        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

    blnSave = True

    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Or State = adStateEditMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    blnSave = False
'    CN.RollbackTrans
'    CN.BeginTrans
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
    If CloseMe = True Then
        Unload Me
    Else
'        txtInvoiceNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid

    CN.BeginTrans
    
    Screen.MousePointer = vbHourglass
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        InitNSD
        'Set the recordset
        rs.Open "SELECT * FROM qry_Receipts WHERE ReceiptID=" & ReceiptPK, CN, adOpenStatic, adLockOptimistic

        dtpDate.Value = Date
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        GeneratePK
        DisplayForAdding
    Else
        'Set the recordset
        rs.Open "SELECT * FROM qry_Sales_Return WHERE SalesReturnID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If State = adStateViewMode Then
            cmdCancel.Caption = "Close"
                   
            DisplayForViewing
        Else
            InitNSD
            DisplayForEditing
        End If
        
    End If
    
    Screen.MousePointer = vbDefault
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Sales_Return")
End Sub

Private Sub ResetEntry()
    'nsdStock.ResetValue
    txtPrice.Tag = 0
    txtPrice.Text = "0.00"
    txtQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If HaveAction = True Then
'        frmSalesReturn.RefreshRecords
'    End If
    
    Set frmSalesReturnAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If State = adStateViewMode Then Exit Sub
        
        dcUnit.Text = ""
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 10), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        nsdStock.Text = .TextMatrix(.RowSel, 2)
        nsdStock.Tag = .TextMatrix(.RowSel, 10) 'Add tag coz boundtext is empty
        txtQty = .TextMatrix(.RowSel, 3)
        
        
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtPrice = toMoney(.TextMatrix(.RowSel, 5))
        txtGross(1) = toMoney(.TextMatrix(.RowSel, 6))
        txtDisc = toMoney(.TextMatrix(.RowSel, 7))
        txtNetAmount = toMoney(.TextMatrix(.RowSel, 8))
        
        If Grid.Rows = 2 And Grid.TextMatrix(1, 10) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtPrice_Change()
    txtQty_Change
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    txtPrice.Text = toMoney(toNumber(txtPrice.Text))
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
    End If
    
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
    
    intQtyOld = txtQty.Text
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Used to Add record
Private Sub DisplayForAdding()
    On Error GoTo err
    nsdClient.Tag = rs!ClientID
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = rs!Company
    txtCity.Text = rs!City
    
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

'Used to edit record
Private Sub DisplayForEditing()
    On Error GoTo err
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = rs!Company
    txtCity.Text = rs!City
    txtReturnSlipNo.Text = rs!ReturnSlipNo
    dtpDate.Value = rs!Date
    cboStatus.Text = rs!Status_Alias
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtNet.Text = toMoney(rs![NetAmount])
    txtNotes.Text = rs![Notes]
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Sales_Return_Detail WHERE SalesReturnID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![Discount] * 100
                    .TextMatrix(1, 8) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 9) = RSDetails![ReturnType]
                    .TextMatrix(1, 10) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![ReturnType]
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                End If
                cIRowCount = cIRowCount + 1
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 9
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    dtpDate.Visible = True
    txtDate.Visible = False

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo err
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = rs!Company
    txtCity.Text = rs!City
    txtReturnSlipNo.Text = rs!ReturnSlipNo
    txtDate.Text = rs![Date]
    cboStatus.Text = rs!Status_Alias
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtNet.Text = toMoney(rs![NetAmount])
    txtNotes.Text = rs![Notes]
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Sales_Return_Detail WHERE SalesReturnID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![Discount] * 100
                    .TextMatrix(1, 8) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 9) = RSDetails![ReturnType]
                    .TextMatrix(1, 10) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![ReturnType]
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 9
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    'Disable commands
    LockInput Me, True

    dtpDate.Visible = False
    txtDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnAdd.Visible = False

    'Resize and reposition the controls
    'Shape3.Top = 4800
    'Label11.Top = 4800
    'Line1(1).Visible = False
    'Line2(1).Visible = False
    Grid.Top = 3460
    Grid.Height = 3050
    
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub txtPrice_GotFocus()
    HLText txtPrice
End Sub

Private Sub InitNSD()
    'For Client
    With nsdClient
        .ClearColumn
        .AddColumn "Client ID", 800
        .AddColumn "Company", 2264.88
        .AddColumn "City", 2670.23
        .AddColumn "Owner's Name", 2670.23
        .AddColumn "Credit Term", 0
        .Connection = CN.ConnectionString
        
        .sqlFields = "ClientID, Company, City, OwnersName, CreditTerm"
        .sqlTables = "qry_Clients1"
        
        .sqlSortOrder = "Company ASC"
        
        .BoundField = "ClientID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Clients Record"
        
    End With
    
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,StockID"
        .sqlTables = "Stocks"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Products"
    End With
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
        .Cols = 11
        .ColSel = 10
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1000
        .ColWidth(2) = 2505
        .ColWidth(3) = 1000
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 1100
        .ColWidth(8) = 1200
        .ColWidth(9) = 1200
        .ColWidth(10) = 0

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Price"
        .TextMatrix(0, 6) = "Gross"
        .TextMatrix(0, 7) = "Discount(%)"
        .TextMatrix(0, 8) = "Net Amount"
        .TextMatrix(0, 9) = "Type of Return"
        .TextMatrix(0, 10) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbLeftJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSSalesReturn As New Recordset

    If State = adStateAddMode Then Exit Sub

    RSSalesReturn.CursorLocation = adUseClient
    RSSalesReturn.Open "SELECT * FROM Sales_Return_Detail WHERE SalesReturnID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSSalesReturn.RecordCount > 0 Then
        RSSalesReturn.MoveFirst
        While Not RSSalesReturn.EOF
            CurrRow = getFlexPos(Grid, 10, RSSalesReturn!StockID)

            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Sales_Return_Detail", "SalesReturnDetailID", "", True, RSSalesReturn!SalesReturnDetailID
                End If
            End With
            RSSalesReturn.MoveNext
        Wend
    End If
End Sub
