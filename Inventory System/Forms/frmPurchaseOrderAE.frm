VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmPurchaseOrderAE 
   BorderStyle     =   0  'None
   ClientHeight    =   8820
   ClientLeft      =   30
   ClientTop       =   -60
   ClientWidth     =   12075
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8820
   ScaleWidth      =   12075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFreightCharges 
      Caption         =   "Freight Charges"
      Height          =   315
      Left            =   7380
      TabIndex        =   65
      Top             =   8250
      Width           =   1515
   End
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Purchase Order Tasks"
      Height          =   315
      Left            =   210
      TabIndex        =   64
      Top             =   8280
      Width           =   2085
   End
   Begin VB.TextBox txtCreditTerm 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   1410
      Width           =   1455
   End
   Begin VB.TextBox txtDeliveryTime 
      Height          =   315
      Left            =   6900
      TabIndex        =   9
      Top             =   2100
      Width           =   1455
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmPurchaseOrderAE.frx":0000
      Left            =   10290
      List            =   "frmPurchaseOrderAE.frx":000A
      TabIndex        =   10
      Text            =   "On Hold"
      Top             =   2100
      Width           =   1515
   End
   Begin VB.TextBox txtPONo 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1770
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox txtGenDiscPerc 
      Height          =   315
      Left            =   1770
      TabIndex        =   4
      Top             =   2130
      Width           =   1605
   End
   Begin VB.TextBox txtContact 
      Height          =   315
      Left            =   1770
      TabIndex        =   3
      Top             =   1770
      Width           =   2325
   End
   Begin VB.TextBox txtSalesman 
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   1410
      Width           =   2325
   End
   Begin VB.TextBox txtLocation 
      Height          =   315
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1050
      Width           =   3315
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   150
      ScaleHeight     =   630
      ScaleWidth      =   11700
      TabIndex        =   33
      Top             =   2730
      Width           =   11700
      Begin VB.TextBox txtDiscPercent 
         Height          =   315
         Left            =   7050
         TabIndex        =   15
         Text            =   "0"
         Top             =   225
         Width           =   735
      End
      Begin VB.TextBox txtExtDiscAmt 
         Height          =   315
         Left            =   8610
         TabIndex        =   17
         Text            =   "0"
         Top             =   225
         Width           =   1005
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   225
         Width           =   1290
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   4470
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   225
         Width           =   1185
      End
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   2775
         TabIndex        =   12
         Text            =   "0"
         Top             =   225
         Width           =   660
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10800
         TabIndex        =   18
         Top             =   255
         Width           =   840
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   315
         Left            =   9660
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   240
         Width           =   1035
      End
      Begin VB.TextBox txtExtDiscPerc 
         Height          =   315
         Left            =   7830
         TabIndex        =   16
         Text            =   "0"
         Top             =   225
         Width           =   735
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   225
         Width           =   2700
         _ExtentX        =   4763
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
         TabIndex        =   13
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
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   5
         Left            =   7050
         TabIndex        =   58
         Top             =   0
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.Amt"
         Height          =   240
         Index           =   3
         Left            =   8670
         TabIndex        =   57
         Top             =   30
         Width           =   1020
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5775
         TabIndex        =   40
         Top             =   0
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   39
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   38
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
         TabIndex        =   37
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   36
         Top             =   0
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9840
         TabIndex        =   35
         Top             =   30
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7800
         TabIndex        =   34
         Top             =   30
         Width           =   840
      End
   End
   Begin VB.TextBox txtDesc 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10410
      Locked          =   -1  'True
      TabIndex        =   26
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6630
      Width           =   1425
   End
   Begin VB.TextBox txtNotes 
      Height          =   1065
      Left            =   135
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Tag             =   "Remarks"
      Top             =   6900
      Width           =   4110
   End
   Begin VB.TextBox txtGross 
      Alignment       =   1  'Right Justify
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
      Left            =   10410
      Locked          =   -1  'True
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6330
      Width           =   1425
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   210
      Picture         =   "frmPurchaseOrderAE.frx":0020
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Remove"
      Top             =   3780
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   10485
      TabIndex        =   21
      Top             =   8250
      Width           =   1335
   End
   Begin VB.TextBox txtNet 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10410
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7680
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   10410
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   6930
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   10410
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   7230
      Visible         =   0   'False
      Width           =   1425
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   150
      TabIndex        =   31
      Top             =   8100
      Width           =   11670
      _ExtentX        =   20585
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2640
      Left            =   150
      TabIndex        =   32
      Top             =   3660
      Width           =   11685
      _ExtentX        =   20611
      _ExtentY        =   4657
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
   Begin ctrlNSDataCombo.NSDataCombo nsdVendor 
      Height          =   315
      Left            =   1770
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
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   9090
      TabIndex        =   20
      Top             =   8250
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   6900
      TabIndex        =   5
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44630019
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   6900
      Locked          =   -1  'True
      TabIndex        =   22
      Top             =   720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   285
      Left            =   6900
      TabIndex        =   8
      Top             =   1770
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44630019
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDeliveryDate 
      Height          =   285
      Left            =   6900
      TabIndex        =   62
      Top             =   1770
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Term"
      Height          =   255
      Left            =   5520
      TabIndex        =   63
      Top             =   1410
      Width           =   1305
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Time"
      Height          =   285
      Left            =   5550
      TabIndex        =   61
      Top             =   2130
      Width           =   1275
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date"
      Height          =   285
      Left            =   5550
      TabIndex        =   60
      Top             =   1770
      Width           =   1275
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   8940
      TabIndex        =   59
      Top             =   2130
      Width           =   1275
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8715
      Left            =   60
      Top             =   60
      Width           =   11955
   End
   Begin VB.Shape Shape1 
      Height          =   8055
      Left            =   120
      Top             =   600
      Width           =   11805
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "General Disc (%)"
      Height          =   255
      Left            =   420
      TabIndex        =   56
      Top             =   2130
      Width           =   1305
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact"
      Height          =   225
      Left            =   180
      TabIndex        =   55
      Top             =   1770
      Width           =   1545
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Local Purchase Details"
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
      TabIndex        =   43
      Top             =   3360
      Width           =   4365
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Purchase Order"
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
      TabIndex        =   41
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Salesman"
      Height          =   225
      Left            =   180
      TabIndex        =   54
      Top             =   1440
      Width           =   1545
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "PO Date"
      Height          =   285
      Index           =   1
      Left            =   5550
      TabIndex        =   53
      Top             =   705
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
      Left            =   180
      TabIndex        =   52
      Top             =   1080
      Width           =   1545
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Location"
      Height          =   285
      Left            =   5550
      TabIndex        =   51
      Top             =   1035
      Width           =   1275
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   180
      TabIndex        =   50
      Top             =   750
      Width           =   1545
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
      Left            =   8310
      TabIndex        =   49
      Top             =   6660
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   150
      TabIndex        =   48
      Top             =   6630
      Width           =   990
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
      Left            =   8310
      TabIndex        =   47
      Top             =   6360
      Width           =   2040
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   180
      X2              =   11850
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   11850
      Y1              =   2580
      Y2              =   2580
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
      Left            =   8310
      TabIndex        =   46
      Top             =   7710
      Width           =   2040
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   150
      Top             =   3360
      Width           =   11700
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
      Left            =   8310
      TabIndex        =   45
      Top             =   6960
      Visible         =   0   'False
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
      Left            =   8310
      TabIndex        =   44
      Top             =   7260
      Visible         =   0   'False
      Width           =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   10080
      X2              =   11790
      Y1              =   7620
      Y2              =   7620
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5010
      TabIndex        =   42
      Top             =   4050
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   11805
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Purchase Order Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "Modification History"
      End
      Begin VB.Menu mnu_ReceiveItem 
         Caption         =   "Receive Items"
      End
      Begin VB.Menu mnu_Vat 
         Caption         =   "Show VAT && Taxbase"
      End
   End
End
Attribute VB_Name = "frmPurchaseOrderAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset 'Main recordset for Invoice

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
    
    CurrRow = getFlexPos(Grid, 13, nsdStock.Tag)
    intStockID = nsdStock.Tag

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 13) = "" Then
                .TextMatrix(1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(1, 2) = nsdStock.Text
                .TextMatrix(1, 3) = txtQty.Text
                .TextMatrix(1, 4) = 0
                .TextMatrix(1, 5) = 0
                .TextMatrix(1, 6) = dcUnit.Text
                .TextMatrix(1, 7) = toMoney(txtPrice.Text)
                .TextMatrix(1, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(1, 9) = toMoney(txtDiscPercent.Text)
                .TextMatrix(1, 10) = toNumber(txtExtDiscPerc.Text)
                .TextMatrix(1, 11) = toMoney(txtExtDiscAmt.Text)
                .TextMatrix(1, 12) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(1, 13) = intStockID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(.Rows - 1, 2) = nsdStock.Text
                .TextMatrix(.Rows - 1, 3) = txtQty.Text
                .TextMatrix(.Rows - 1, 4) = 0
                .TextMatrix(.Rows - 1, 5) = 0
                .TextMatrix(.Rows - 1, 6) = dcUnit.Text
                .TextMatrix(.Rows - 1, 7) = toMoney(txtPrice.Text)
                .TextMatrix(.Rows - 1, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(.Rows - 1, 9) = toMoney(txtDiscPercent.Text)
                .TextMatrix(.Rows - 1, 10) = toNumber(txtExtDiscPerc.Text)
                .TextMatrix(.Rows - 1, 11) = toMoney(txtExtDiscAmt.Text)
                .TextMatrix(.Rows - 1, 12) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(.Rows - 1, 13) = intStockID
                
                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 8))
                txtGross(2).Text = Format$(cIGross, "#,##0.00")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
                txtNet.Text = Format$(cIAmount, "#,##0.00")
                
                'Compute discount
                curDiscPerc = .TextMatrix(1, 8) * .TextMatrix(1, 9) / 100
                curExtDiscPerc = .TextMatrix(1, 8) * .TextMatrix(1, 10) / 100
                
                cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
                
                txtDesc.Text = Format$(cDAmount, "#,##0.00")
                
                .TextMatrix(CurrRow, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(CurrRow, 2) = nsdStock.Text
                .TextMatrix(CurrRow, 3) = txtQty.Text
                .TextMatrix(CurrRow, 4) = 0
                .TextMatrix(CurrRow, 5) = 0
                .TextMatrix(CurrRow, 6) = dcUnit.Text
                .TextMatrix(CurrRow, 7) = toMoney(txtPrice.Text)
                .TextMatrix(CurrRow, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(CurrRow, 9) = toMoney(txtDiscPercent.Text)
                .TextMatrix(CurrRow, 10) = toNumber(txtExtDiscPerc.Text)
                .TextMatrix(CurrRow, 11) = toMoney(txtExtDiscAmt.Text)
                .TextMatrix(CurrRow, 12) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(CurrRow, 13) = intStockID

            Else
                Exit Sub
            End If
        End If
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
        
        'Highlight the current row's column
        .ColSel = 12
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
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 8))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        curDiscPerc = .TextMatrix(1, 8) * .TextMatrix(1, 9) / 100
        curExtDiscPerc = .TextMatrix(1, 8) * .TextMatrix(1, 10) / 100
        
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

Private Sub cmdFreightCharges_Click()
    With frmFreightCharges
        .POID = PK
        .VendorPK = nsdVendor.Tag
        
        .show 1
    End With
End Sub

Private Sub mnu_ReceiveItem_Click()
    Dim RSDetails As New Recordset
    
    RSDetails.CursorLocation = adUseClient
    
    If Right(txtLocation.Text, 5) = "Local" Then 'check if local purchase
        RSDetails.CursorLocation = adUseClient
        RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
        
        If RSDetails.RecordCount > 0 Then
            With frmPOReceiveLocalAE
                .State = adStateAddMode
                .PK = PK
                .show vbModal
            End With
        Else
          MsgBox "All items are already delivered to VT.", vbInformation
        End If
    Else
        RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
        
        If RSDetails.RecordCount > 0 Then
            With frmForwardersGuideAE
                .State = adStateAddMode
                .PK = PK
                .show vbModal
            End With
        Else
            MsgBox "All items are already forwarded.", vbInformation
        End If
    End If
End Sub

Private Sub CmdTasks_Click()
    PopupMenu mnu_Tasks
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSStocks As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSStocks.CursorLocation = adUseClient
    RSStocks.Open "SELECT * FROM Purchase_Order_Detail WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSStocks.RecordCount > 0 Then
        RSStocks.MoveFirst
        While Not RSStocks.EOF
            CurrRow = getFlexPos(Grid, 13, RSStocks!StockID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Purchase_Order_Detail", "PODetailID", "", True, RSStocks!PODetailID
                End If
            End With
            RSStocks.MoveNext
        Wend
    End If
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

Private Sub nsdVendor_Change()
    If nsdVendor.DisableDropdown = False Then
        txtLocation.Text = nsdVendor.getSelValueAt(3)
        txtGenDiscPerc.Text = nsdVendor.getSelValueAt(4)
        txtCreditTerm.Text = nsdVendor.getSelValueAt(5)
    End If
    
    nsdStock.sqlwCondition = "VendorID = " & nsdVendor.BoundText
End Sub

Private Sub txtContact_GotFocus()
    HLText txtContact
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

Private Sub txtExtDiscAmt_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtDiscAmt_Validate(Cancel As Boolean)
    txtExtDiscAmt.Text = toMoney(toNumber(txtExtDiscAmt.Text))
End Sub

Private Sub txtExtDiscPerc_Change()
    ComputeGrossNet
End Sub

Private Sub txtExtDiscPerc_Click()
    txtQty_Change
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub txtExtDiscPerc_GotFocus()
    HLText txtExtDiscPerc
End Sub

Private Sub txtExtDiscPerc_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtDiscPerc_Validate(Cancel As Boolean)
    txtExtDiscPerc.Text = toNumber(txtExtDiscPerc.Text)
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If nsdVendor.Text = "" Then
        MsgBox "Please select a vendor.", vbExclamation
        nsdVendor.SetFocus
        Exit Sub
    End If
    
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before you can save this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
       
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    DeleteItems
    
    'Save the record
    With RS
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![POID] = PK
            ![VendorID] = nsdVendor.BoundText
            ![PONo] = txtPONo.Text
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        Else
            RS.Close
            RS.Open "SELECT * FROM Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
            
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        
        ![Date] = dtpDate.Value
        ![Contact] = txtContact.Text
        ![DeliveryDate] = dtpDeliveryDate.Value
        ![DeliveryTime] = txtDeliveryTime.Text
        ![Status] = IIf(cboStatus.Text = "Ordered", True, False)
        ![Notes] = txtNotes.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)

        ![Status] = IIf(cboStatus.Text = "Ordered", True, False)
        ![Notes] = txtNotes.Text
        
        
        .Update
    End With
   
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Purchase_Order_Detail WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
   
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSDetails.AddNew

                RSDetails![POID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 13))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 6))
                RSDetails![Price] = toNumber(.TextMatrix(c, 7))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 10)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 11))
                
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 13))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 6))
                RSDetails![Price] = toNumber(.TextMatrix(c, 7))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 10)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 11))
                
                RSDetails.Update
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
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
            txtPONo.Text = Format(PK, "0000000000")
         Else
            Unload Me
        End If
    Else
        MsgBox "Changes in record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        nsdVendor.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    InitGrid
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        InitNSD
        
        'Set the recordset
         RS.Open "SELECT * FROM Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
         dtpDate.Value = Date
         txtDeliveryTime.Text = Time()
         mnu_History.Enabled = False
         mnu_ReceiveItem.Enabled = False
         cmdFreightCharges.Visible = False
         GeneratePK
         txtPONo.Text = Format(PK, "0000000000")
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        RS.Open "SELECT * FROM qry_Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If State = adStateViewMode Then
            cmdCancel.Caption = "Close"
            mnu_History.Enabled = True
                   
            DisplayForViewing
        Else
            InitNSD
            DisplayForEditing
            
            mnu_ReceiveItem.Enabled = False
            
            nsdStock.sqlwCondition = "VendorID = " & RS!VendorID
        End If
        
        If ForCusAcc = True Then
            Me.Icon = frmPurchaseOrder.Icon
        End If

        Screen.MousePointer = vbDefault
    End If
    'Initialize Graphics
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Purchase_Order")
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
        .ColSel = 12
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 2430
        .ColWidth(3) = 465
        .ColWidth(4) = 900
        .ColWidth(5) = 700
        .ColWidth(6) = 1000
        .ColWidth(7) = 1005
        .ColWidth(8) = 900
        .ColWidth(9) = 690
        .ColWidth(10) = 995
        .ColWidth(11) = 1150
        .ColWidth(12) = 1000
        .ColWidth(13) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Item"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Qty Rcvd"
        .TextMatrix(0, 5) = "Qty Due"
        .TextMatrix(0, 6) = "Unit"
        .TextMatrix(0, 7) = "Price" 'Supplier Price
        .TextMatrix(0, 8) = "Gross"
        .TextMatrix(0, 9) = "Disc(%)"
        .TextMatrix(0, 10) = "Ext. Disc(%)"
        .TextMatrix(0, 11) = "Ext. Disc(Amt)"
        .TextMatrix(0, 12) = "Net Amount"
        .TextMatrix(0, 13) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbLeftJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
'        .ColAlignment(9) = vbRightJustify
'        .ColAlignment(10) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtQty.Text = "0"
    txtPrice.Tag = 0
    txtPrice.Text = "0.00"
    txtDiscPercent.Text = "0"
    txtExtDiscPerc.Text = "0"
    txtExtDiscAmt.Text = "0"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmPurchaseOrder.RefreshRecords
    End If
    
    Set frmPurchaseOrderAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If State = adStateViewMode Then Exit Sub
        
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 13), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        nsdStock.Text = .TextMatrix(.RowSel, 2)
        nsdStock.Tag = .TextMatrix(.RowSel, 13) 'Add tag coz boundtext is empty
        txtQty.Text = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 6)
        txtPrice.Text = toMoney(.TextMatrix(.RowSel, 7))
        txtGross(1).Text = toMoney(.TextMatrix(.RowSel, 8))
        txtDiscPercent.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtExtDiscPerc.Text = toMoney(.TextMatrix(.RowSel, 10))
        txtExtDiscAmt.Text = toMoney(.TextMatrix(.RowSel, 11))
        txtNetAmount.Text = toMoney(.TextMatrix(.RowSel, 12))
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 13) = "" Then '13 = StockID
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

Private Sub nsdStock_Change()
    On Error Resume Next
    
    txtQty.Text = "0"
    
    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText & " ORDER BY qry_Unit.Order ASC", "Unit", dcUnit, "UnitID", True
    
    nsdStock.Tag = nsdStock.BoundText
    
    txtPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Supplier Price
    txtExtDiscPerc.Text = toMoney(nsdStock.getSelValueAt(5)) 'Ext Disc (%)
    txtExtDiscAmt.Text = toMoney(nsdStock.getSelValueAt(6)) 'Ext Disc (Amt)
      
    If toNumber(nsdStock.getSelValueAt(4)) = 0 Then
      txtDiscPercent.Text = toNumber(txtGenDiscPerc.Text)
    Else
      txtDiscPercent.Text = toNumber(nsdStock.getSelValueAt(4))
    End If
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtExtDiscAmt_Change()
    ComputeGrossNet
End Sub

Private Sub txtExtDiscAmt_GotFocus()
  HLText txtExtDiscAmt
End Sub

Private Sub txtGenDiscPerc_Change()
  ComputeGrossNet
End Sub

Private Sub txtGenDiscPerc_GotFocus()
  HLText txtGenDiscPerc
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtPrice_Change()
  ComputeGrossNet
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    txtPrice.Text = toMoney(toNumber(txtPrice.Text))
End Sub

Private Sub txtQty_Change()
    ComputeGrossNet
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
        btnAdd.Enabled = False
    Else
        btnAdd.Enabled = True
    End If
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry
    
    nsdVendor.Text = ""
    txtPONo.Text = ""
    txtLocation.Text = ""
    txtDate.Text = ""
    txtContact.Text = ""
    txtSalesman.Text = ""
    'txtShippingInstructions.Text = ""
    'txtAdditionalInstructions.Text = ""
    'txtDeclaredAs.Text = ""
    'txtDeclaredValue.Text = ""
    
    txtGross(2).Text = "0.00"
    txtDesc.Text = "0.00"
    txtTaxBase.Text = "0.00"
    txtVat.Text = "0.00"
    txtNet.Text = "0.00"

    cIAmount = 0
    cDAmount = 0

    nsdVendor.SetFocus
End Sub

'Used to display record
Private Sub DisplayForEditing()
    On Error GoTo err
    nsdVendor.Tag = RS![VendorID]
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = RS!Company
    txtCreditTerm.Text = RS!CreditTerm
    txtPONo.Text = RS!PONo
    txtLocation.Text = RS!Location
    dtpDate.Value = Format(RS![Date], "MMM-dd-yy")
    txtContact.Text = RS![Contact]
    dtpDeliveryDate.Value = RS![DeliveryDate]
    txtDeliveryTime.Text = RS![DeliveryTime]
    txtSalesman.Text = RS![Salesman]
    cboStatus.Text = RS!Status_Alias
        
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = RS![Notes]
    txtGenDiscPerc.Text = toNumber(RS![GenDiscPercent])
    
    cIGross = toNumber(txtGross(2).Text)
    cDAmount = toNumber(txtDesc.Text)
    cIAmount = toNumber(txtNet.Text)
    cIRowCount = 0
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 13) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![QtyReceived]
                    .TextMatrix(1, 5) = RSDetails![QtyDue]
                    .TextMatrix(1, 6) = RSDetails![Unit]
                    .TextMatrix(1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 12) = toMoney(RSDetails!NetAmount)
                    .TextMatrix(1, 13) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![QtyReceived]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails!NetAmount)
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
    nsdVendor.Tag = RS![VendorID]
    nsdVendor.DisableDropdown = True
    nsdVendor.TextReadOnly = True
    nsdVendor.Text = RS!Company
    txtCreditTerm.Text = RS!CreditTerm
    txtPONo.Text = RS!PONo
    txtLocation.Text = RS!Location
    txtDate.Text = Format(RS![Date], "MMM-dd-yy")
    txtContact.Text = RS![Contact]
    txtDeliveryDate.Text = RS![DeliveryDate]
    txtDeliveryTime.Text = RS![DeliveryTime]
    txtSalesman.Text = RS![Salesman]
    cboStatus.Text = RS!Status_Alias
        
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = RS![Notes]
    txtGenDiscPerc.Text = toNumber(RS![GenDiscPercent])
    
'    cIRowCount = 0
    
    'Display the details
    Dim RSDetails As New Recordset
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
'          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 13) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![QtyReceived]
                    .TextMatrix(1, 5) = RSDetails![QtyDue]
                    .TextMatrix(1, 6) = RSDetails![Unit]
                    .TextMatrix(1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 12) = toMoney(RSDetails!NetAmount)
                    .TextMatrix(1, 13) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![QtyReceived]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails!NetAmount)
                    .TextMatrix(.Rows - 1, 13) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 12
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    'Disable commands
    LockInput Me, True
    
    'unlock Notes
    txtNotes.Locked = False
    
    dtpDate.Visible = False
    txtDate.Visible = True
    dtpDeliveryDate.Visible = False
    txtDeliveryDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnAdd.Enabled = False
    
    mnu_ReceiveItem.Visible = True
    mnu_ReceiveItem.Enabled = True
    
    'Resize and reposition the controls
    Shape3.Top = 2500
    Label11.Top = 2500
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 2850
    Grid.Height = 3420

    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

Private Sub InitNSD()
    'For Vendor
    With nsdVendor
        .ClearColumn
        .AddColumn "Supplier ID", 1794.89
        .AddColumn "Supplier", 2264.88
        .AddColumn "Location", 2670.23
        .AddColumn "Gen Disc (%)", 0
        .AddColumn "Credit Term", 0
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "VendorID, Company, Location, GenDiscPercent, CreditTerm"
        .sqlTables = "qry_Vendors1"
        .sqlSortOrder = "Company ASC"
        
        .BoundField = "VendorID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Vendors Record"
    End With
    
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Product", 4085.26
        .AddColumn "Supplier Price", 1500
        .AddColumn "Disc (%)", 1500
        .AddColumn "Ext Disc (%)", 0
        .AddColumn "Ext Disc (Amt)", 0
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,SupplierPrice,DiscPercent,ExtDiscPercent,ExtDiscAmount,StockID"
        '.sqlTables = "Stocks"
        .sqlTables = "qry_Vendors_Stocks"

        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Stocks"
    End With
End Sub

Private Sub txtPrice_GotFocus()
    HLText txtPrice
End Sub

Private Sub txtSalesman_GotFocus()
    HLText txtSalesman
End Sub
