VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmPOReceiveLocalAE 
   BorderStyle     =   0  'None
   ClientHeight    =   8985
   ClientLeft      =   -30
   ClientTop       =   -60
   ClientWidth     =   12915
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   8985
   ScaleWidth      =   12915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Purchase Receive Tasks"
      Height          =   315
      Left            =   270
      TabIndex        =   60
      Top             =   8460
      Width           =   2085
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmPurchaseOrderReceiveAE.frx":0000
      Left            =   5880
      List            =   "frmPurchaseOrderReceiveAE.frx":000A
      TabIndex        =   6
      Text            =   " "
      Top             =   1530
      Width           =   2325
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   240
      ScaleHeight     =   630
      ScaleWidth      =   12510
      TabIndex        =   46
      Top             =   3360
      Width           =   12510
      Begin MSDataListLib.DataCombo dcWarehouse 
         Height          =   315
         Left            =   10770
         TabIndex        =   15
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   225
         Width           =   2715
      End
      Begin VB.TextBox txtExtDiscPerc 
         Height          =   315
         Left            =   7830
         TabIndex        =   13
         Text            =   "0"
         Top             =   225
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   315
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   48
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   240
         Width           =   1035
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11670
         TabIndex        =   16
         Top             =   255
         Width           =   780
      End
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   2775
         TabIndex        =   9
         Text            =   "0"
         Top             =   225
         Width           =   660
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   4470
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   47
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox txtExtDiscAmt 
         Height          =   315
         Left            =   8610
         TabIndex        =   14
         Text            =   "0"
         Top             =   225
         Width           =   1005
      End
      Begin VB.TextBox txtDiscPercent 
         Height          =   315
         Left            =   7050
         TabIndex        =   12
         Text            =   "0"
         Top             =   225
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   225
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Label6 
         Caption         =   "Warehouse"
         Height          =   255
         Left            =   10770
         TabIndex        =   59
         Top             =   30
         Width           =   855
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7800
         TabIndex        =   57
         Top             =   30
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9840
         TabIndex        =   56
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   55
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
         Left            =   0
         TabIndex        =   54
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   53
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   52
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   51
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.Amt"
         Height          =   240
         Index           =   3
         Left            =   8670
         TabIndex        =   50
         Top             =   30
         Width           =   1020
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   5
         Left            =   7050
         TabIndex        =   49
         Top             =   0
         Width           =   840
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Remarks"
      Height          =   1335
      Left            =   10650
      TabIndex        =   7
      Top             =   1590
      Width           =   2025
      Begin VB.OptionButton optLost 
         Caption         =   "Lost"
         Height          =   195
         Left            =   210
         TabIndex        =   45
         Top             =   810
         Width           =   1365
      End
      Begin VB.OptionButton optReceived 
         Caption         =   "Received"
         Height          =   195
         Left            =   210
         TabIndex        =   44
         Top             =   420
         Value           =   -1  'True
         Width           =   1365
      End
   End
   Begin VB.TextBox txtDRDate 
      Height          =   314
      Left            =   1530
      TabIndex        =   3
      Top             =   1890
      Width           =   1905
   End
   Begin VB.TextBox txtDRNo 
      Height          =   314
      Left            =   1530
      TabIndex        =   2
      Top             =   1530
      Width           =   1905
   End
   Begin VB.TextBox txtDeliveryNo 
      Height          =   314
      Left            =   1530
      TabIndex        =   4
      Top             =   2250
      Width           =   1905
   End
   Begin VB.TextBox txtPONo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1530
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   690
      Width           =   3315
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7980
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   11325
      TabIndex        =   19
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   9960
      TabIndex        =   18
      Top             =   8490
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   300
      Picture         =   "frmPurchaseOrderReceiveAE.frx":0021
      Style           =   1  'Graphical
      TabIndex        =   24
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
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   23
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6630
      Width           =   1425
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   255
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "Remarks"
      Top             =   6870
      Width           =   5910
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   11250
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7530
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   11250
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   7230
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   20
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6930
      Width           =   1425
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   26
      Top             =   8340
      Width           =   12330
      _ExtentX        =   21749
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2490
      Left            =   240
      TabIndex        =   27
      Top             =   4020
      Width           =   12495
      _ExtentX        =   22040
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
      AllowUserResizing=   1
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
   Begin ctrlNSDataCombo.NSDataCombo nsdVendor 
      Height          =   315
      Left            =   1530
      TabIndex        =   1
      Top             =   1050
      Width           =   3690
      _ExtentX        =   6509
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
   Begin MSComCtl2.DTPicker dtpDRDate 
      Height          =   315
      Left            =   1530
      TabIndex        =   39
      Top             =   1890
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   12517379
      CurrentDate     =   38989
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   315
      Left            =   1530
      TabIndex        =   5
      Top             =   2610
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   12517379
      CurrentDate     =   38989
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   4530
      TabIndex        =   58
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive item details"
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
      TabIndex        =   28
      Top             =   3150
      Width           =   4365
   End
   Begin VB.Label Label18 
      Caption         =   "DR No."
      Height          =   225
      Left            =   210
      TabIndex        =   43
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label Label17 
      Caption         =   "DR Date"
      Height          =   225
      Left            =   210
      TabIndex        =   42
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label Label7 
      Caption         =   "Delivery No."
      Height          =   225
      Left            =   210
      TabIndex        =   41
      Top             =   2280
      Width           =   1275
   End
   Begin VB.Label Label12 
      Caption         =   "Delivery Date"
      Height          =   225
      Left            =   210
      TabIndex        =   40
      Top             =   2640
      Width           =   1275
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   180
      X2              =   12690
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   180
      X2              =   12660
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   210
      TabIndex        =   38
      Top             =   690
      Width           =   1275
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   37
      Top             =   1050
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Receive Items"
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
      Left            =   270
      TabIndex        =   32
      Top             =   150
      Width           =   4905
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   120
      Width           =   12645
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5130
      TabIndex        =   36
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   10170
      X2              =   11880
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   240
      Top             =   3150
      Width           =   12480
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
      Left            =   9150
      TabIndex        =   35
      Top             =   8010
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   240
      X2              =   12690
      Y1              =   3090
      Y2              =   3090
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   240
      X2              =   12690
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
      Left            =   9150
      TabIndex        =   34
      Top             =   6660
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   270
      TabIndex        =   33
      Top             =   6600
      Width           =   990
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8895
      Left            =   60
      Top             =   60
      Width           =   12825
   End
   Begin VB.Shape Shape1 
      Height          =   8235
      Left            =   150
      Top             =   630
      Width           =   12645
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
      Left            =   9150
      TabIndex        =   31
      Top             =   7560
      Visible         =   0   'False
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
      Left            =   9150
      TabIndex        =   30
      Top             =   7260
      Visible         =   0   'False
      Width           =   2040
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
      Left            =   9150
      TabIndex        =   29
      Top             =   6960
      Width           =   2040
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Purchase Receive Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "Modification History"
      End
      Begin VB.Menu mnu_ReturnItems 
         Caption         =   "Return Items"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnu_Vat 
         Caption         =   "Show VAT && Taxbase"
      End
   End
End
Attribute VB_Name = "frmPOReceiveLocalAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit (PO)
Public InvoicePK            As Long 'Variable used to get what record is going to edit (Invoice)
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset 'Main recordset for Invoice
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Allowed value for receive qty

Private Sub btnUpdate_Click()
    Dim CurrRow As Integer
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency
    Dim intQty As Integer
    
    CurrRow = getFlexPos(Grid, 12, txtStock.Tag)

    'Add to grid
    With Grid
        .Row = CurrRow
        
        'Restore back the invoice amount and discount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 6))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 10))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        
        'Compute discount
        curDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 7) / 100
        curExtDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 8) / 100
        
        cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        
        .TextMatrix(CurrRow, 3) = txtQty.Text
        .TextMatrix(CurrRow, 4) = dcUnit.Text
        .TextMatrix(CurrRow, 5) = toMoney(txtPrice.Text)
        .TextMatrix(CurrRow, 6) = toMoney(txtGross(1).Text)
        .TextMatrix(CurrRow, 7) = toMoney(txtDiscPercent.Text)
        .TextMatrix(CurrRow, 8) = toNumber(txtExtDiscPerc.Text)
        .TextMatrix(CurrRow, 9) = toMoney(txtExtDiscAmt.Text)
        .TextMatrix(CurrRow, 10) = toMoney(toNumber(txtNetAmount.Text))
        .TextMatrix(CurrRow, 11) = toMoney(dcWarehouse.Text)
        
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        
        'Compute discount
        curDiscPerc = txtGross(1).Text * txtDiscPercent.Text / 100
        curExtDiscPerc = txtGross(1).Text * txtExtDiscPerc.Text / 100
        
        cDAmount = curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text
        
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        'Save to Purchase Order Details
        Dim RSPODetails As New Recordset
        
        RSPODetails.CursorLocation = adUseClient
        RSPODetails.Open "SELECT * From Purchase_Order_Detail Where POID = " & txtPONo.Tag, CN, adOpenStatic, adLockOptimistic
        
        'add qty received in Purchase Order Details
        RSPODetails.Find "[StockID] = " & txtStock.Tag, , adSearchForward, 1
       
        If txtQty > intQtyOld Then
            intQty = txtQty.Text - intQtyOld
            RSPODetails!QtyReceived = toNumber(RSPODetails!QtyReceived) + intQty
        Else
            intQty = intQtyOld - txtQty
            RSPODetails!QtyReceived = toNumber(RSPODetails!QtyReceived) - intQty
        End If
        
        RSPODetails.Update
        '-----------------
        
        'Highlight the current row's column
        .ColSel = 10
        'Display a remove button
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
End Sub

Private Sub btnRemove_Click()
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency
    
    'Remove selected load product
    With Grid
        'Update grooss to current purchase amount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 6))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 10))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        curDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 7) / 100
        curExtDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 8) / 100
        
        cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If txtDRNo = "" Then
        MsgBox "Please don't leave DR No Field blank", vbInformation
        txtDRNo.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to return before saving this record.", vbExclamation
        Exit Sub
    End If
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Purchase_Order_Receive
    Dim RSReceive As New Recordset

    RSReceive.CursorLocation = adUseClient
    RSReceive.Open "SELECT * FROM Purchase_Order_Receive_Local WHERE InvoiceID=" & InvoicePK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    DeleteItems
    
    'Save the record
    With RSReceive
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            
            InvoicePK = getIndex("Purchase_Order_Receive_Local")
            ![InvoiceID] = InvoicePK
            ![POID] = PK
          
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        End If
        ![DRNo] = txtDRNo.Text
        ![DRDate] = dtpDRDate.Value
        ![DeliveryNo] = txtDeliveryNo.Text
        ![DeliveryDate] = dtpDeliveryDate.Value
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
        
        ![Notes] = txtNotes.Text
        ![Remarks] = IIf(optReceived.Value, "R", "L") 'R = Received; L = Lost
        ![Status] = IIf(cboStatus.Text = "Received", True, False)
        
        ![DateModified] = Now
        ![LastUserFK] = CurrUser.USER_PK
        
        .Update
    End With
   
    'Connection for Purchase_Order_Receiving_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Purchase_Order_Receive_Local_Detail WHERE InvoiceID=" & InvoicePK, CN, adOpenStatic, adLockOptimistic
    
    'Save to stock card
    Dim RSStockCard As New Recordset

    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic
    
    'Add qty ordered to qty onhand
    Dim RSStockUnit As New Recordset

    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * From Stock_Unit", CN, adOpenStatic, adLockOptimistic
       
'    'Save to Purchase Order Details
'    Dim RSPODetails As New Recordset
'
'    RSPODetails.CursorLocation = adUseClient
'    RSPODetails.Open "SELECT * From Purchase_Order_Detail Where POID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    'Save to Landed Cost table
    Dim RSLandedCost As New Recordset

    RSLandedCost.CursorLocation = adUseClient
    
    With Grid
        
        'Save the details of the records to Purchase_Order_Receive_Local_Detail
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
            
                RSDetails.AddNew
    
                RSDetails![InvoiceID] = InvoicePK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 12))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
                RSDetails![WarehouseID] = getValueAt("SELECT WarehouseID,Warehouse FROM Warehouses WHERE Warehouse='" & .TextMatrix(c, 11) & "'", "WarehouseID")
    
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 12))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
                RSDetails![WarehouseID] = getValueAt("SELECT WarehouseID,Warehouse FROM Warehouses WHERE Warehouse='" & .TextMatrix(c, 11) & "'", "WarehouseID")

                RSDetails.Update
                
            End If
            
            If State = adStateAddMode Then
                RSStockCard.AddNew
    
                RSStockCard!Type = "PRL"
                RSStockCard!UnitID = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSStockCard!RefNo1 = txtDRNo.Text
                RSStockCard!Incoming = toNumber(.TextMatrix(c, 3))
                RSStockCard!Cost = toNumber(.TextMatrix(c, 5))
                RSStockCard!StockID = toNumber(.TextMatrix(c, 12))
    
                RSStockCard.Update
                '-----------------
                
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 12)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")

                RSStockUnit!Incoming = RSStockUnit!Incoming + toNumber(.TextMatrix(c, 3))
                RSStockUnit.Update
            ElseIf cboStatus.Text = "On Hold" And State = adStateEditMode Then
                RSStockCard.Filter = "StockID = " & toNumber(.TextMatrix(c, 12)) & " AND RefNo1 = '" & txtDRNo.Text & "'"
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 12)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
               
                'Restore Pending from incoming
                RSStockUnit!Incoming = RSStockUnit!Incoming - RSStockCard!Incoming
                
                RSStockUnit.Update
                '-----------------
                
                'Update Incoming, Overight Incoming encoded in add mode
                RSStockCard!Incoming = toNumber(.TextMatrix(c, 3))
    
                RSStockCard.Update
                '-----------------

                RSStockUnit!Incoming = RSStockUnit!Incoming + toNumber(.TextMatrix(c, 3))
                
                RSStockUnit.Update
            End If
            
            If cboStatus.Text = "Received" And optReceived.Value = True Then
                'Add qty received to stock card
                RSStockCard.Filter = "StockID = " & toNumber(.TextMatrix(c, 12)) & " AND RefNo1 = '" & txtDRNo.Text & "'"

                RSStockCard!Incoming = RSStockCard!Incoming - toNumber(.TextMatrix(c, 3))
                RSStockCard!Pieces1 = toNumber(.TextMatrix(c, 3)) 'RSStockCard!Pieces1 + toNumber(.TextMatrix(c, 3))

                RSStockCard.Update
                '-----------------

                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 12)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                
                RSStockUnit!Incoming = RSStockUnit!Incoming - toNumber(.TextMatrix(c, 3))
                RSStockUnit!Onhand = RSStockUnit!Onhand + toNumber(.TextMatrix(c, 3))
                
                RSStockUnit.Update
                '-----------------
                                    
'                'add qty received in Purchase Order Details
'                RSPODetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 12)), , adSearchForward, 1
'                RSPODetails!QtyReceived = toNumber(RSPODetails!QtyReceived) + toNumber(.TextMatrix(c, 3))
'
'                RSPODetails.Update
'                '-----------------
            
                RSLandedCost.Open "SELECT * FROM Landed_Cost WHERE StockID=" & toNumber(.TextMatrix(c, 12)) & " ORDER BY LandedCostID DESC", CN, adOpenStatic, adLockOptimistic
                
                If RSLandedCost.RecordCount > 0 Then
                    If RSLandedCost!SupplierPrice <> toNumber(.TextMatrix(c, 5)) Then
AddLandedCost:
                        RSLandedCost.AddNew
        
                        RSLandedCost!VendorID = getValueAt("SELECT VendorID,Company FROM Vendors WHERE Company='" & nsdVendor.Text & "'", "VendorID")
                        RSLandedCost!RefNo = txtDRNo.Text
                        RSLandedCost!StockID = .TextMatrix(c, 12)
                        RSLandedCost!Unit = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                        RSLandedCost!Date = Date
                        RSLandedCost!SupplierPrice = toNumber(.TextMatrix(c, 5))
                        RSLandedCost!Discount = toNumber(.TextMatrix(c, 7)) / 100
                        RSLandedCost!ExtDiscPercent = toNumber(.TextMatrix(c, 8)) / 100
                        RSLandedCost!ExtDiscAmount = toNumber(.TextMatrix(c, 9))
                        
                        RSLandedCost.Update
                    End If
                ElseIf RSLandedCost.RecordCount = 0 Then
                    GoSub AddLandedCost
                End If
                
                RSLandedCost.Close
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

Private Sub CmdTasks_Click()
    PopupMenu mnu_Tasks
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        txtDRNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid

    CN.BeginTrans

    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "UnitID", True
    bind_dc "SELECT * FROM Warehouses", "Warehouse", dcWarehouse, "WarehouseID", True

    Screen.MousePointer = vbHourglass
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDRDate.Value = Date
        dtpDeliveryDate.Value = Date
        mnu_History.Enabled = False
                   
        DisplayForAdding
    ElseIf State = adStateEditMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Purchase_Order_Receive_Local WHERE InvoiceID=" & InvoicePK, CN, adOpenStatic, adLockOptimistic
        
        dtpDRDate.Value = Date
        dtpDeliveryDate.Value = Date
        mnu_History.Enabled = False
                   
        DisplayForEditing
    Else
        'Set the recordset
        RS.Open "SELECT * FROM qry_Purchase_Order_Receive_Local WHERE InvoiceID=" & InvoicePK, CN, adOpenStatic, adLockOptimistic
        
        cmdCancel.Caption = "Close"
        DisplayForViewing
    End If
    
    Screen.MousePointer = vbDefault
    
    'Initialize Graphics
    With MAIN
        'cmdGenerate.Picture = .i16x16.ListImages(14).Picture
        'cmdNew.Picture = .i16x16.ListImages(10).Picture
        'cmdReset.Picture = .i16x16.ListImages(15).Picture
    End With
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Local_Purchase")
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
        .Cols = 13
        .ColSel = 11
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 2430
        .ColWidth(3) = 465
        .ColWidth(4) = 1000
        .ColWidth(5) = 1005
        .ColWidth(6) = 900
        .ColWidth(7) = 690
        .ColWidth(8) = 995
        .ColWidth(9) = 1150
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Item"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Price" 'Supplier Price
        .TextMatrix(0, 6) = "Gross"
        .TextMatrix(0, 7) = "Disc(%)"
        .TextMatrix(0, 8) = "Ext. Disc(%)"
        .TextMatrix(0, 9) = "Ext. Disc(Amt)"
        .TextMatrix(0, 10) = "Net Amount"
        .TextMatrix(0, 11) = "Warehouse"
        .TextMatrix(0, 12) = "Stock ID"
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
        .ColAlignment(10) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    txtStock.Text = ""
    txtQty.Text = "0"
    txtPrice.Tag = 0
    txtPrice.Text = "0.00"
    txtDiscPercent.Text = "0"
    txtExtDiscPerc.Text = "0"
    txtExtDiscAmt.Text = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If HaveAction = True Then
    '    frmLocalPurchaseReturn.RefreshRecords
    'End If
    
    Set frmPOReceiveLocalAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtStock.Text = .TextMatrix(.RowSel, 2)
        txtStock.Tag = .TextMatrix(.RowSel, 12) 'Create tag to get the StockID
        intQtyOld = IIf(.TextMatrix(.RowSel, 3) = "", 0, .TextMatrix(.RowSel, 3))
        txtQty = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtPrice = toMoney(.TextMatrix(.RowSel, 5))
        txtGross(1) = toMoney(.TextMatrix(.RowSel, 6))
        txtDiscPercent.Text = toMoney(.TextMatrix(.RowSel, 7))
        txtExtDiscPerc.Text = toMoney(.TextMatrix(.RowSel, 8))
        txtExtDiscAmt.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtNetAmount = toMoney(.TextMatrix(.RowSel, 10))
        
        If State = adStateViewMode Then Exit Sub
        If Grid.Rows = 2 And Grid.TextMatrix(1, 12) = "" Then
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

Private Sub mnu_History_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub

Private Sub mnu_ReturnItems_Click()
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Receive_Local_Detail WHERE InvoiceID=" & InvoicePK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        With frmPOReturnAE
            .State = adStateAddMode
            .PK = InvoicePK
            .show vbModal
        End With
    Else
        MsgBox "All items are already returned.", vbInformation
    End If
End Sub

Private Sub mnu_Vat_Click()
  If mnu_Vat.Caption = "Show VAT && Taxbase" Then
    Label5.Visible = True
    Label8.Visible = True
    txtTaxBase.Visible = True
    txtVat.Visible = True
    mnu_Vat.Caption = "Hide VAT && Taxbase"
  Else
    Label5.Visible = False
    Label8.Visible = False
    txtTaxBase.Visible = False
    txtVat.Visible = False
    mnu_Vat.Caption = "Show VAT && Taxbase"
  End If
End Sub

Private Sub txtDeliveryNo_GotFocus()
    HLText txtDeliveryNo
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtDiscPercent_Change()
    ComputeGrossNet
End Sub

Private Sub txtDiscPercent_GotFocus()
    HLText txtDiscPercent
End Sub

Private Sub txtDiscPercent_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtDRNo_GotFocus()
    HLText txtDRNo
End Sub

Private Sub txtExtDiscAmt_Change()
    ComputeGrossNet
End Sub

Private Sub txtExtDiscAmt_GotFocus()
    HLText txtExtDiscAmt
End Sub

Private Sub txtExtDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtDiscAmt_Validate(Cancel As Boolean)
    txtExtDiscAmt.Text = toMoney(toNumber(txtExtDiscAmt.Text))
End Sub

Private Sub txtExtDiscPerc_Change()
    ComputeGrossNet
End Sub

Private Sub txtExtDiscPerc_GotFocus()
    HLText txtExtDiscPerc
End Sub

Private Sub txtExtDiscPerc_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_LostFocus()
    Dim intQtyDue As Integer
      
    intQtyDue = getValueAt("SELECT QtyDue FROM qry_Purchase_Order_Detail WHERE POID=" & PK, "QtyDue")
    If txtQty.Text > (intQtyDue + intQtyOld) Then
        MsgBox "Overdelivery for " & txtStock.Text & ".", vbInformation
        txtQty.Text = intQtyOld
    End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtPrice_Change()
'    txtQty_Change
    ComputeGrossNet
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    txtPrice.Text = toMoney(toNumber(txtPrice.Text))
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnUpdate.Enabled = False
    Else
        btnUpdate.Enabled = True
    End If
    
    ComputeGrossNet
'    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)))
'    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * toNumber(txtPrice.Text)) - ((toNumber(txtDiscPercent.Text) / 100) * toNumber(toNumber(txtQty.Text) * toNumber(txtPrice.Text))))
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
    
'    intQtyOld = txtQty.Text
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Used to edit record
Private Sub DisplayForAdding()
    On Error GoTo err
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = RS!Company
    txtPONo.Tag = RS!POID
    txtPONo.Text = RS!PONo
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = RS![Notes]
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    
    'Save to Purchase Order Details
'    Dim RSPODetails As New Recordset
'
'    RSPODetails.CursorLocation = adUseClient
'    RSPODetails.Open "SELECT * From Purchase_Order_Detail Where POID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 12) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![QtyDue]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = "MV1"
                    .TextMatrix(1, 12) = RSDetails![StockID]
                    
                    'add qty received in Purchase Order Details
                    RSDetails!QtyReceived = toNumber(RSDetails!QtyReceived) + toNumber(RSDetails![QtyDue])
                    
                    RSDetails.Update
                    '-----------------
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![Price]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = "MV1"
                    .TextMatrix(.Rows - 1, 12) = RSDetails![StockID]
                    
                    'add qty received in Purchase Order Details
                    RSDetails!QtyReceived = toNumber(RSDetails!QtyReceived) + toNumber(RSDetails![QtyDue])
                    
                    RSDetails.Update
                    '-----------------
                End If
                cIRowCount = cIRowCount + 1
            End With
            RSDetails.MoveNext
        Wend
        
        Grid.Row = 1
        Grid.ColSel = 11
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    dtpDRDate.Visible = True
    txtDRDate.Visible = False
    lblStatus.Visible = False
    cboStatus.Visible = False
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
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = RS!Company
    txtPONo.Tag = RS!POID
    txtPONo.Text = RS!PONo
    PK = RS!POID 'get POID to be make a reference for QtyDue, etc.
    txtDRNo.Text = RS!DRNo
    
    txtDRNo.Text = RS![DRNo]
    dtpDRDate.Value = RS![DRDate]
    txtDeliveryNo.Text = RS![DeliveryNo]
    dtpDeliveryDate.Value = RS![DeliveryDate]
    txtNotes.Text = RS![Notes]
    optReceived.Value = IIf(RS![Remarks] = "R", True, False)
    txtNotes.Text = IIf(IsNull(RS![Notes]), "", RS![Notes])
    cboStatus.Text = RS!Status_Alias
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    cIRowCount = 0
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Receive_Local_Detail WHERE InvoiceID=" & InvoicePK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 12) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(1, 12) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(.Rows - 1, 12) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 11
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    dtpDRDate.Visible = True
    txtDRDate.Visible = False
    lblStatus.Visible = True
    cboStatus.Visible = True

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
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = RS!Company
    txtPONo.Text = RS!PONo
    txtDRNo.Text = RS!DRNo
    
    txtDRNo.Text = RS![DRNo]
    dtpDRDate.Value = RS![DRDate]
    txtDeliveryNo.Text = RS![DeliveryNo]
    dtpDeliveryDate.Value = RS![DeliveryDate]
    txtNotes.Text = RS![Notes]
    optReceived.Value = IIf(RS![Remarks] = "R", True, False)
    cboStatus.Text = RS!Status_Alias
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = IIf(IsNull(RS![Notes]), "", RS![Notes])
       
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Receive_Local_Detail WHERE InvoiceID=" & InvoicePK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 12) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(1, 12) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(.Rows - 1, 12) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 10
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

    dtpDRDate.Visible = True
    txtDRDate.Visible = False
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnUpdate.Visible = False

'    CmdReturn.Left = cmdSave.Left
'    CmdReturn.Visible = True
    
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

Private Sub ComputeGrossNet()
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency
    
    If toNumber(txtQty.Text) < 1 Or toMoney(txtPrice.Text) < 1 Then Exit Sub
    
    txtGross(1).Text = toMoney(toNumber(txtQty.Text) * toNumber(txtPrice.Text))
    
    curDiscPerc = txtGross(1).Text * txtDiscPercent.Text / 100
    curExtDiscPerc = txtGross(1).Text * txtExtDiscPerc.Text / 100
    
    txtNetAmount.Text = txtGross(1).Text - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)

    If toNumber(txtQty.Text) < 1 Then
        btnUpdate.Enabled = False
    Else
        btnUpdate.Enabled = True
    End If
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSStocks As New Recordset
'    Dim intStockID As Integer
    
    If State = adStateAddMode Then Exit Sub
    
    RSStocks.CursorLocation = adUseClient
    RSStocks.Open "SELECT * FROM Purchase_Order_Receive_Local_Detail WHERE InvoiceID=" & InvoicePK, CN, adOpenStatic, adLockOptimistic
    If RSStocks.RecordCount > 0 Then
        RSStocks.MoveFirst
        While Not RSStocks.EOF
            CurrRow = getFlexPos(Grid, 12, RSStocks!StockID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Purchase_Order_Receive_Local_Detail", "PORecDetailID", "", True, RSStocks!PORecDetailID
                    
                    If cboStatus.Text = "Received" And optReceived.Value = True Then getStockCardID RSStocks!StockID, txtDRNo.Text
                    'DelRecwSQL "Purchase_Order_Receive_Local_Detail", "PORecDetailID", "", True, RSStocks!PORecDetailID
                End If
            End With
            RSStocks.MoveNext
        Wend
    End If
End Sub

'delete record from Stock Card
Private Sub getStockCardID(StockID As Integer, strRefNo As String)
    Dim RSStockCard As New Recordset
    
    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card WHERE StockID=" & StockID & " AND RefNo = '" & strRefNo & "'", CN, adOpenStatic, adLockOptimistic
    If RSStockCard.RecordCount > 0 Then
        RSStockCard.Delete
        
        RSStockCard.Update
    End If
End Sub
