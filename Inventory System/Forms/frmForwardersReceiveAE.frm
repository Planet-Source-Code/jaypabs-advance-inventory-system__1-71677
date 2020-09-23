VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmForwardersReceiveAE 
   BorderStyle     =   0  'None
   ClientHeight    =   8730
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   12885
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8730
   ScaleWidth      =   12885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtReceiptDate 
      Height          =   315
      Left            =   6240
      TabIndex        =   64
      Top             =   2340
      Width           =   1785
   End
   Begin VB.TextBox txtRefNo 
      Height          =   314
      Left            =   9060
      TabIndex        =   61
      Top             =   1980
      Width           =   2415
   End
   Begin VB.ComboBox cboRef 
      Height          =   315
      ItemData        =   "frmForwardersReceiveAE.frx":0000
      Left            =   9060
      List            =   "frmForwardersReceiveAE.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   60
      Top             =   1590
      Width           =   2415
   End
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Purchase Receive Tasks"
      Height          =   315
      Left            =   240
      TabIndex        =   59
      Top             =   8160
      Width           =   2085
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   210
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Tag             =   "Remarks"
      Top             =   6570
      Width           =   5910
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   210
      ScaleHeight     =   630
      ScaleWidth      =   12495
      TabIndex        =   26
      Top             =   3090
      Width           =   12495
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   30
         TabIndex        =   8
         Top             =   255
         Width           =   2715
      End
      Begin VB.TextBox txtExtDiscPerc 
         Height          =   315
         Left            =   7830
         TabIndex        =   13
         Text            =   "0"
         Top             =   255
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   315
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   28
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11640
         TabIndex        =   15
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   2775
         TabIndex        =   9
         Text            =   "0"
         Top             =   255
         Width           =   660
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   4470
         TabIndex        =   11
         Text            =   "0.00"
         Top             =   255
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   27
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   255
         Width           =   1290
      End
      Begin VB.TextBox txtExtDiscAmt 
         Height          =   315
         Left            =   8610
         TabIndex        =   14
         Text            =   "0"
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtDiscPercent 
         Height          =   315
         Left            =   7050
         TabIndex        =   12
         Text            =   "0"
         Top             =   255
         Width           =   735
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   10
         Top             =   255
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcWarehouse 
         Height          =   315
         Left            =   10740
         TabIndex        =   57
         Top             =   270
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.Label Label15 
         Caption         =   "Warehouse"
         Height          =   255
         Left            =   10740
         TabIndex        =   58
         Top             =   60
         Width           =   855
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7800
         TabIndex        =   37
         Top             =   60
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   225
         Left            =   9690
         TabIndex        =   36
         Top             =   60
         Width           =   915
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   35
         Top             =   30
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
         TabIndex        =   34
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   33
         Top             =   30
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   32
         Top             =   30
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   31
         Top             =   30
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.Amt"
         Height          =   240
         Index           =   3
         Left            =   8670
         TabIndex        =   30
         Top             =   60
         Width           =   1020
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   5
         Left            =   7050
         TabIndex        =   29
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "Return Items"
      Height          =   315
      Left            =   8610
      TabIndex        =   18
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
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
      Top             =   7650
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   11370
      TabIndex        =   20
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   9990
      TabIndex        =   19
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmForwardersReceiveAE.frx":008F
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Remove"
      Top             =   3900
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
      TabIndex        =   24
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6300
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   11250
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   11250
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   6900
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   11250
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtShippingGuideNo 
      Height          =   314
      Left            =   1860
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1590
      Width           =   1545
   End
   Begin VB.TextBox txtPONo 
      Height          =   285
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1905
   End
   Begin VB.TextBox txtSupplier 
      Height          =   314
      Left            =   1830
      TabIndex        =   1
      Top             =   990
      Width           =   3075
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmForwardersReceiveAE.frx":0241
      Left            =   9990
      List            =   "frmForwardersReceiveAE.frx":024B
      TabIndex        =   2
      Text            =   " "
      Top             =   630
      Width           =   2325
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      ItemData        =   "frmForwardersReceiveAE.frx":0262
      Left            =   1860
      List            =   "frmForwardersReceiveAE.frx":026C
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2340
      Width           =   3225
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   210
      TabIndex        =   38
      Top             =   8010
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2490
      Left            =   210
      TabIndex        =   39
      Top             =   3720
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
   Begin ctrlNSDataCombo.NSDataCombo nsdShippingCo 
      Height          =   315
      Left            =   1860
      TabIndex        =   4
      Top             =   1950
      Width           =   3240
      _ExtentX        =   5715
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
   Begin ctrlNSDataCombo.NSDataCombo nsdLocal 
      Height          =   315
      Left            =   6240
      TabIndex        =   6
      Top             =   1590
      Width           =   1800
      _ExtentX        =   3175
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
   Begin VB.TextBox txtDeliveryDate 
      Height          =   315
      Left            =   6240
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1950
      Width           =   1785
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Receipt Date"
      Height          =   225
      Left            =   5130
      TabIndex        =   65
      Top             =   2370
      Width           =   1065
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ref No."
      Height          =   255
      Left            =   8160
      TabIndex        =   63
      Top             =   2010
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Reference"
      Height          =   285
      Left            =   8160
      TabIndex        =   62
      Top             =   1590
      Width           =   825
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   6
      Left            =   225
      TabIndex        =   56
      Top             =   6300
      Width           =   990
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
      Left            =   210
      TabIndex        =   55
      Top             =   2850
      Width           =   4365
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5130
      TabIndex        =   54
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   10920
      X2              =   12630
      Y1              =   7590
      Y2              =   7590
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
      TabIndex        =   53
      Top             =   7680
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
      Left            =   9150
      TabIndex        =   52
      Top             =   6330
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
      Left            =   9150
      TabIndex        =   51
      Top             =   7230
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
      TabIndex        =   50
      Top             =   6930
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
      TabIndex        =   49
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Local"
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
      Index           =   11
      Left            =   5370
      TabIndex        =   48
      Top             =   1620
      Width           =   825
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Guide No."
      Height          =   255
      Left            =   270
      TabIndex        =   47
      Top             =   1620
      Width           =   1560
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      Caption         =   "Class"
      Height          =   255
      Left            =   270
      TabIndex        =   46
      Top             =   2340
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Receiving History"
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
      Left            =   210
      TabIndex        =   45
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Company"
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
      Left            =   270
      TabIndex        =   44
      Top             =   1980
      Width           =   1560
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   540
      TabIndex        =   43
      Top             =   690
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   210
      X2              =   12690
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   210
      X2              =   12690
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   225
      Left            =   5370
      TabIndex        =   42
      Top             =   1980
      Width           =   825
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier"
      Height          =   225
      Left            =   540
      TabIndex        =   41
      Top             =   1050
      Width           =   1245
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   12750
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   12720
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   8640
      TabIndex        =   40
      Top             =   660
      Width           =   1305
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8625
      Left            =   60
      Top             =   60
      Width           =   12795
   End
   Begin VB.Shape Shape1 
      Height          =   7995
      Left            =   150
      Top             =   600
      Width           =   12615
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   120
      Width           =   12615
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2820
      Width           =   12495
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Purchase Receive Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "Modification History"
      End
      Begin VB.Menu mnu_Return 
         Caption         =   "Return Items"
      End
      Begin VB.Menu mnu_Vat 
         Caption         =   "Show VAT && Taxbase"
      End
   End
End
Attribute VB_Name = "frmForwardersReceiveAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit (PO)
Public ForwarderReceivePK   As Long 'Variable used to get what record is going to edit (Invoice)
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset 'Main recordset for Invoice
Dim intQtyOld               As Integer 'Allowed value for receive qty
Dim cLocalTrucking          As Currency
Dim cSidewalkHandling       As Currency
Dim intVendorID             As Integer
Dim blnSave                 As Boolean

Private Sub btnUpdate_Click()
    Dim CurrRow As Integer
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency
    Dim intQty As Integer

    If cboClass.Text = "" Or nsdLocal.Text = "" Then
        MsgBox "Class & Local Forwarder Fields needs input", vbInformation
        Exit Sub
    End If
    
    CurrRow = getFlexPos(Grid, 13, txtStock.Tag)

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
        .TextMatrix(CurrRow, 12) = dcWarehouse.Text
                
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
        
        'Save to Forwarder Details
        Dim RSForwarderDetails As New Recordset
    
        RSForwarderDetails.CursorLocation = adUseClient
        RSForwarderDetails.Open "SELECT * From Forwarders_Detail Where ForwarderID = " & PK, CN, adOpenStatic, adLockOptimistic
        
        'add qty received in Purchase Order Details
        RSForwarderDetails.Find "[StockID] = " & txtStock.Tag, , adSearchForward, 1
       
        If txtQty > intQtyOld Then
            intQty = txtQty.Text - intQtyOld
            RSForwarderDetails!QtyReceived = toNumber(RSForwarderDetails!QtyReceived) + intQty
        Else
            intQty = intQtyOld - txtQty
            RSForwarderDetails!QtyReceived = toNumber(RSForwarderDetails!QtyReceived) - intQty
        End If
        
        RSForwarderDetails.Update
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
        curDiscPerc = .TextMatrix(.RowSel, 6) * .TextMatrix(.RowSel, 7) / 100
        curExtDiscPerc = .TextMatrix(.RowSel, 6) * .TextMatrix(.RowSel, 8) / 100
        
        cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        Dim RSDetails As New Recordset
    
        RSDetails.CursorLocation = adUseClient
        RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & PK & " AND StockID = " & .TextMatrix(.RowSel, 13), CN, adOpenStatic, adLockOptimistic
        If RSDetails.RecordCount > 0 Then
            'restore back qty that was previously added in DisplayForAdding procedure
            RSDetails!QtyReceived = toNumber(RSDetails!QtyReceived) - .TextMatrix(.RowSel, 3)
            
            RSDetails.Update
            '-----------------
        End If
        
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

Private Sub mnu_Return_Click()
    Dim RSPOReturn As New Recordset

    RSPOReturn.CursorLocation = adUseClient
    RSPOReturn.Open "SELECT POReturnID FROM Purchase_Order_Return WHERE RefNo=" & ForwarderReceivePK, CN, adOpenStatic, adLockOptimistic
    
    With frmPOReturnAE
        If RSPOReturn.RecordCount > 0 Then 'if record exist then edit record
            Dim blnStatus As Boolean
            
            blnStatus = getValueAt("SELECT POReturnID,Status FROM Purchase_Order_Return WHERE POReturnID=" & RSPOReturn!POReturnID, "Status")
            
            If blnStatus Then 'true
                .State = adStateViewMode
            Else
                .State = adStateEditMode
            End If
            
            .PK = RSPOReturn!POReturnID
        Else
            .State = adStateAddMode
            .ReceivePK = ForwarderReceivePK
        End If
        
        .show vbModal
    End With
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If cIRowCount < 1 Then
        MsgBox "Please enter item to return before saving this record.", vbExclamation
        Exit Sub
    End If
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Purchase_Order_Receive
    Dim RSReceive As New Recordset

    RSReceive.CursorLocation = adUseClient
    RSReceive.Open "SELECT * FROM Forwarders_Receive WHERE ForwarderReceiveID=" & ForwarderReceivePK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    DeleteItems
    
    'Save the record
    With RSReceive
        If State = adStateAddMode Then
    
            .AddNew
            
            ForwarderReceivePK = getIndex("Forwarders_Receive")
            ![ForwarderReceiveID] = ForwarderReceivePK
            ![ForwarderID] = PK
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        End If
        
        ![ShippingCompanyID] = IIf(nsdShippingCo.BoundText = "", nsdShippingCo.Tag, nsdShippingCo.BoundText)
        ![ShippingGuideNo] = txtShippingGuideNo.Text
'        ![Ship] = txtShip.Text
        ![Class] = cboClass.ListIndex
        ![LocalForwarderID] = IIf(nsdLocal.BoundText = "", nsdLocal.Tag, nsdLocal.BoundText)
        ![DeliveryDate] = txtDeliveryDate.Text
        ![ReceiptDate] = txtReceiptDate.Text
        ![Ref] = cboRef.Text
        ![RefNo] = txtRefNo.Text
'        ![DRNo] = txtDRNo.Text
'        ![BLNo] = txtBLNo.Text
'        ![TruckNo] = txtTruckNo.Text
'        ![VanNo] = txtVanNo.Text
'        ![VoyageNo] = txtVoyageNo.Text
        ![Status] = IIf(cboStatus.Text = "Received", True, False)
        ![Notes] = txtNotes.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
    
        ![DateModified] = Now
        ![LastUserFK] = CurrUser.USER_PK
                
        .Update
    End With
  
    'Connection for Forwarders_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Forwarders_Receive_Detail WHERE ForwarderReceiveID=" & ForwarderReceivePK, CN, adOpenStatic, adLockOptimistic
          
    'Save to stock card
    Dim RSStockCard As New Recordset

    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic
          
    'Add qty ordered to qty onhand
    Dim RSStockUnit As New Recordset

    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * From Stock_Unit", CN, adOpenStatic, adLockOptimistic
          
    'Save to Purchase Order Details
'    Dim RSForwarderDetails As New Recordset
'
'    RSForwarderDetails.CursorLocation = adUseClient
'    RSForwarderDetails.Open "SELECT * From Forwarders_Detail Where ForwarderID = " & PK, CN, adOpenStatic, adLockOptimistic
    
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
    
                RSDetails![ForwarderReceiveID] = ForwarderReceivePK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 13))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
                RSDetails![WarehouseID] = getValueAt("SELECT WarehouseID,Warehouse FROM Warehouses WHERE Warehouse='" & .TextMatrix(c, 12) & "'", "WarehouseID")
    
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 13))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
                RSDetails![WarehouseID] = getValueAt("SELECT WarehouseID,Warehouse FROM Warehouses WHERE Warehouse='" & .TextMatrix(c, 12) & "'", "WarehouseID")
    
                RSDetails.Update
            End If
            
            If State = adStateAddMode Then
                RSStockCard.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND RefNo1 = '" & txtRefNo.Text & "'"
                
                If RSStockCard.RecordCount = 0 Then
                    'Add to stock card
                    RSStockCard.AddNew
    
                    RSStockCard!Type = "PR"
                    RSStockCard!UnitID = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                    RSStockCard!RefNo1 = txtRefNo.Text
                    RSStockCard!Incoming = toNumber(.TextMatrix(c, 3))
                    RSStockCard!Cost = toNumber(.TextMatrix(c, 5))
                    RSStockCard!StockID = toNumber(.TextMatrix(c, 13))
                Else
                    RSStockCard!Incoming = RSStockCard!Incoming + toNumber(.TextMatrix(c, 3))
                End If
                
                RSStockCard.Update
                '-----------------
                
                'Deduct pending and add incoming in Stock Unit
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")

                RSStockUnit!Pending = RSStockUnit!Pending - toNumber(.TextMatrix(c, 3))
                RSStockUnit!Incoming = RSStockUnit!Incoming + toNumber(.TextMatrix(c, 3))
                RSStockUnit.Update
            ElseIf cboStatus.Text = "On Hold" And State = adStateEditMode Then
                RSStockCard.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND RefNo1 = '" & txtRefNo.Text & "'"
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
               
                'Restore Pending from incoming
                RSStockUnit!Pending = RSStockUnit!Pending + RSStockCard!Incoming
                RSStockUnit!Incoming = RSStockUnit!Incoming - RSStockCard!Incoming
                
                RSStockUnit.Update
                '-----------------
                
                'Update Incoming, Overight Incoming encoded in add mode
                RSStockCard!Incoming = toNumber(.TextMatrix(c, 3))
    
                RSStockCard.Update
                '-----------------

                'Deduct pending and add incoming in Stock Unit
                RSStockUnit!Pending = RSStockUnit!Pending - toNumber(.TextMatrix(c, 3))
                RSStockUnit!Incoming = RSStockUnit!Incoming + toNumber(.TextMatrix(c, 3))
                
                RSStockUnit.Update
            End If
            
            If cboStatus.Text = "Received" Then
                RSStockCard.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND RefNo1 = '" & txtRefNo.Text & "'"

                RSStockCard!Incoming = RSStockCard!Incoming - toNumber(.TextMatrix(c, 3))
                RSStockCard!Pieces1 = RSStockCard!Pieces1 + toNumber(.TextMatrix(c, 3))

                RSStockCard.Update
                '-----------------

                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 13)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                
                RSStockUnit!Incoming = RSStockUnit!Incoming - toNumber(.TextMatrix(c, 3))
                RSStockUnit!Onhand = RSStockUnit!Onhand + toNumber(.TextMatrix(c, 3))
                
                RSStockUnit.Update
                '-----------------
            
                'add qty received in Forwarders Details
'                RSForwarderDetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 13)), , adSearchForward, 1
'                RSForwarderDetails!QtyReceived = toNumber(RSForwarderDetails!QtyReceived) + toNumber(.TextMatrix(c, 3))
'
'                RSForwarderDetails.Update
                
                RSLandedCost.Open "SELECT * FROM Landed_Cost WHERE StockID=" & toNumber(.TextMatrix(c, 13)) & " ORDER BY LandedCostID DESC", CN, adOpenStatic, adLockOptimistic
                
                If RSLandedCost.RecordCount > 0 Then
                    If toNumber(RSLandedCost!SupplierPrice) <> toNumber(.TextMatrix(c, 5)) Then
AddLandedCost:
                        'Save to landed cost table
                        RSLandedCost.AddNew
        
                        RSLandedCost!VendorID = intVendorID
                        RSLandedCost!RefNo = txtRefNo.Text
                        RSLandedCost!StockID = .TextMatrix(c, 13)
                        RSLandedCost!Unit = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                        RSLandedCost!Date = Date
                        RSLandedCost!SupplierPrice = toNumber(.TextMatrix(c, 5))
                        RSLandedCost!Discount = toNumber(.TextMatrix(c, 7)) / 100
                        RSLandedCost!ExtDiscPercent = toNumber(.TextMatrix(c, 8)) / 100
                        RSLandedCost!ExtDiscAmount = toNumber(.TextMatrix(c, 9))
                        RSLandedCost!Freight = toNumber(.TextMatrix(c, 11))
                        
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
        txtShippingGuideNo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid
    InitNSD
        
    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "UnitID", True
    bind_dc "SELECT * FROM Warehouses", "Warehouse", dcWarehouse, "WarehouseID", True
    
    Screen.MousePointer = vbHourglass
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Forwarders WHERE ForwarderID=" & PK, CN, adOpenStatic, adLockOptimistic
        mnu_Return.Enabled = False
        
        CN.BeginTrans
                   
        DisplayForAdding
    ElseIf State = adStateEditMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Forwarders_Receive WHERE ForwarderReceiveID=" & ForwarderReceivePK, CN, adOpenStatic, adLockOptimistic
        
        mnu_Return.Enabled = False

        CN.BeginTrans

        DisplayForEditing
    Else
        'Set the recordset
        RS.Open "SELECT * FROM qry_Forwarders_Receive WHERE ForwarderReceiveID=" & ForwarderReceivePK, CN, adOpenStatic, adLockOptimistic
        
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
    
    Set frmForwardersReceiveAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateViewMode Then Exit Sub
    
    With Grid
        txtStock.Text = .TextMatrix(.RowSel, 2)
        txtStock.Tag = .TextMatrix(.RowSel, 13) 'Create tag to get the StockID
        intQtyOld = IIf(.TextMatrix(.RowSel, 3) = "", 0, .TextMatrix(.RowSel, 3))
        txtQty = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtPrice = toMoney(.TextMatrix(.RowSel, 5))
        txtGross(1) = toMoney(.TextMatrix(.RowSel, 6))
        txtDiscPercent.Text = toMoney(.TextMatrix(.RowSel, 7))
        txtExtDiscPerc.Text = toMoney(.TextMatrix(.RowSel, 8))
        txtExtDiscAmt.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtNetAmount.Text = toMoney(.TextMatrix(.RowSel, 10))
        dcWarehouse.Text = .TextMatrix(.RowSel, 12)
        
        If State = adStateViewMode Then Exit Sub
        If Grid.Rows = 2 And Grid.TextMatrix(1, 13) = "" Then
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

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtDiscPercent_Change()
'    ComputeGrossNet
End Sub

Private Sub txtDiscPercent_GotFocus()
    HLText txtDiscPercent
End Sub

Private Sub txtDiscPercent_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtDiscAmt_Change()
'    ComputeGrossNet
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
'    ComputeGrossNet
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
      
    intQtyDue = getValueAt("SELECT QtyDue FROM qry_Forwarders_Detail WHERE ForwarderID=" & PK, "QtyDue")
    If txtQty.Text > (intQtyDue + intQtyOld) Then
        MsgBox "Overdelivery for " & txtStock.Text & ".", vbInformation
        txtQty.Text = intQtyOld
    End If
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtPrice_Change()
'    ComputeGrossNet
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
    
'    ComputeGrossNet
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
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo

    PK = RS!ForwarderID 'get ForwarderID to be make a reference for QtyDue, etc.
    nsdShippingCo.Tag = RS![ShippingCompanyID]
    nsdShippingCo.Text = RS![ShippingCompany]
    txtShippingGuideNo.Text = RS![ShippingGuideNo]
'    txtShip.Text = rs![Ship]
    cboClass.ListIndex = RS![Class]
    nsdLocal.Tag = RS![LocalForwarderID]
    nsdLocal.Text = RS![LocalForwarder]
    txtDeliveryDate.Text = RS![DeliveryDate]
    txtReceiptDate.Text = RS![ReceiptDate]
    cboRef.Text = RS![Ref]
    txtRefNo.Text = RS![RefNo]
'    txtDRNo.Text = rs![DRNo]
'    txtBLNo.Text = rs![BLNo]
'    txtTruckNo.Text = rs![TruckNo]
'    txtVanNo.Text = rs![VanNo]
'    txtVoyageNo.Text = rs![VoyageNo]

    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])

    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 13) = "" Then
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
                    .TextMatrix(1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(1, 12) = "MV1"
                    .TextMatrix(1, 13) = RSDetails![StockID]
                    
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
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(.Rows - 1, 12) = "MV1"
                    .TextMatrix(.Rows - 1, 13) = RSDetails![StockID]
                    
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
        Grid.ColSel = 12
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

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
    intVendorID = RS!VendorID
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo
    
    PK = RS!ForwarderID 'get ForwarderID to be make a reference for QtyDue, etc.
    nsdShippingCo.Tag = RS![ShippingCompanyID]
    nsdShippingCo.Text = RS![ShippingCompany]
    txtShippingGuideNo.Text = RS![ShippingGuideNo]
'    txtShip.Text = rs![Ship]
    cboClass.ListIndex = RS![Class]
    nsdLocal.Tag = RS![LocalForwarderID]
    nsdLocal.Text = RS![LocalForwarder]
    txtDeliveryDate.Text = RS![DeliveryDate]
    txtReceiptDate.Text = RS![ReceiptDate]
    cboRef.Text = RS![Ref]
    txtRefNo.Text = RS![RefNo]
'    txtDRNo.Text = rs![DRNo]
'    txtBLNo.Text = rs![BLNo]
'    txtTruckNo.Text = rs![TruckNo]
'    txtVanNo.Text = rs![VanNo]
'    txtVoyageNo.Text = rs![VoyageNo]
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
    RSDetails.Open "SELECT * FROM qry_Forwarders_Receive_Detail WHERE ForwarderReceiveID=" & ForwarderReceivePK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 13) = "" Then
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
                    .TextMatrix(1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(1, 12) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(1, 13) = RSDetails![StockID]
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
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(.Rows - 1, 13) = RSDetails![StockID]
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
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo
    PK = RS!POID 'get POID to be make a reference for QtyDue, etc.
    
    nsdShippingCo.Tag = RS![ShippingCompanyID]
    nsdShippingCo.Text = RS![ShippingCompany]
    txtShippingGuideNo.Text = RS![ShippingGuideNo]
'    txtShip.Text = rs![Ship]
    cboClass.ListIndex = RS![Class]
    nsdLocal.Tag = RS![LocalForwarderID]
    nsdLocal.Text = RS![LocalForwarder]
    txtDeliveryDate.Text = RS![DeliveryDate]
    txtReceiptDate.Text = RS![ReceiptDate]
    cboRef.Text = RS![Ref]
    txtRefNo.Text = RS![RefNo]
'    txtDRNo.Text = rs![DRNo]
'    txtBLNo.Text = rs![BLNo]
'    txtTruckNo.Text = rs![TruckNo]
'    txtVanNo.Text = rs![VanNo]
'    txtVoyageNo.Text = rs![VoyageNo]
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
    RSDetails.Open "SELECT * FROM qry_Forwarders_Receive_Detail WHERE ForwarderReceiveID=" & ForwarderReceivePK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 13) = "" Then
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
                    .TextMatrix(1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(1, 12) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(1, 13) = RSDetails![StockID]
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
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![CostperPackage])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![Warehouse])
                    .TextMatrix(.Rows - 1, 13) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 12
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
    Grid.Top = 3160
    Grid.Height = 3100
    
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

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 14
        .ColSel = 13
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 2430
        .ColWidth(3) = 465
        .ColWidth(4) = 1000
        .ColWidth(5) = 690
        .ColWidth(6) = 995
        .ColWidth(7) = 1150
        .ColWidth(8) = 1000
        .ColWidth(9) = 1150
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 0
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
        .TextMatrix(0, 11) = "Cost Per Package"
        .TextMatrix(0, 12) = "Warehouse"
        .TextMatrix(0, 13) = "Stock ID"
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
        .ColAlignment(10) = vbRightJustify
    End With
End Sub

Private Sub InitNSD()
    'For Shipping Company
    With nsdShippingCo
      .ClearColumn
      .AddColumn "ID", 500.89
      .AddColumn "Shipping Company", 1794.89
      .Connection = CN.ConnectionString
      
      .sqlFields = "ShippingCompanyID, ShippingCompany"
      .sqlTables = "qry_Shipping_Company"
      .sqlSortOrder = "ShippingCompany ASC"
      
      .BoundField = "ShippingCompanyID"
      .PageBy = 25
      .DisplayCol = 2
      
      .setDropWindowSize 7000, 4000
      .TextReadOnly = True
      .SetDropDownTitle = "Shipping Companies"
    End With
    
    With nsdLocal
      .ClearColumn
      .AddColumn "ID", 500.89
      .AddColumn "Local Forwarder", 1794.89
      .Connection = CN.ConnectionString
      
      .sqlFields = "LocalForwarderID, LocalForwarder"
      .sqlTables = "Local_Forwarder"
      .sqlSortOrder = "LocalForwarder ASC"
      
      .BoundField = "LocalForwarderID"
      .PageBy = 25
      .DisplayCol = 2
      
      .setDropWindowSize 7000, 4000
      .TextReadOnly = True
      .SetDropDownTitle = "Local Forwarder"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSStocks As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSStocks.CursorLocation = adUseClient
    RSStocks.Open "SELECT * FROM Forwarders_Receive_Detail WHERE ForwarderReceiveID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSStocks.RecordCount > 0 Then
        RSStocks.MoveFirst
        While Not RSStocks.EOF
            CurrRow = getFlexPos(Grid, 13, RSStocks!StockID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Forwarders_Receive_Detail", "ForwarderReceiveDetailID", "", True, RSStocks!ForwarderReceiveDetailID
                End If
            End With
            RSStocks.MoveNext
        Wend
    End If
End Sub

