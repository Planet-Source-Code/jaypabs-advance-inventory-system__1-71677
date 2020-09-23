VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmSalesReceiptsAE 
   BorderStyle     =   0  'None
   ClientHeight    =   9015
   ClientLeft      =   1635
   ClientTop       =   690
   ClientWidth     =   13725
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9015
   ScaleWidth      =   13725
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAddCharges 
      Enabled         =   0   'False
      Height          =   195
      Left            =   7080
      TabIndex        =   70
      Top             =   750
      Width           =   195
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   315
      Left            =   7080
      TabIndex        =   8
      Top             =   1710
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89653251
      CurrentDate     =   39299
   End
   Begin VB.TextBox txtDeliveryDate 
      Height          =   315
      Left            =   7080
      TabIndex        =   66
      Top             =   1710
      Width           =   2505
   End
   Begin MSDataListLib.DataCombo dcRoute 
      Height          =   315
      Left            =   1590
      TabIndex        =   0
      Top             =   720
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   315
      Left            =   10650
      TabIndex        =   22
      Top             =   8460
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox txtCreditTerm 
      Height          =   285
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2460
      Width           =   1575
   End
   Begin VB.ComboBox cboDeducted 
      Height          =   315
      ItemData        =   "frmReceiptsAE.frx":0000
      Left            =   7080
      List            =   "frmReceiptsAE.frx":000A
      TabIndex        =   10
      Text            =   "No"
      Top             =   2430
      Width           =   2325
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmReceiptsAE.frx":0017
      Left            =   7080
      List            =   "frmReceiptsAE.frx":0021
      TabIndex        =   9
      Text            =   "On Hold"
      Top             =   2070
      Width           =   2325
   End
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Sales Receipts Tasks"
      Height          =   315
      Left            =   210
      TabIndex        =   27
      Top             =   8490
      Width           =   1755
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   12060
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   7500
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   12060
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   7200
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12060
      Locked          =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7950
      Width           =   1425
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   12675
      TabIndex        =   24
      Top             =   8460
      Width           =   795
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   11700
      TabIndex        =   23
      Top             =   8460
      Width           =   855
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmReceiptsAE.frx":0034
      Style           =   1  'Graphical
      TabIndex        =   25
      ToolTipText     =   "Remove"
      Top             =   3930
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
      Left            =   12060
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   240
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Tag             =   "Remarks"
      Top             =   6870
      Width           =   5910
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   12060
      Locked          =   -1  'True
      TabIndex        =   38
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
      ScaleWidth      =   13245
      TabIndex        =   28
      Top             =   2910
      Width           =   13245
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   12690
         TabIndex        =   21
         Top             =   300
         Width           =   510
      End
      Begin VB.CheckBox ckFree 
         Height          =   225
         Left            =   12420
         TabIndex        =   20
         Top             =   330
         Width           =   240
      End
      Begin VB.TextBox txtAddCharges 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   10440
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   300
         Width           =   1065
      End
      Begin VB.TextBox txtCreditTerm2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   11550
         TabIndex        =   19
         Text            =   "0"
         Top             =   300
         Width           =   825
      End
      Begin VB.TextBox txtExtPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5580
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   300
         Width           =   1125
      End
      Begin VB.TextBox txtDisc 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8670
         TabIndex        =   17
         Text            =   "0"
         Top             =   300
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E6FFFF&
         Height          =   285
         Left            =   9480
         Locked          =   -1  'True
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   300
         Width           =   855
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   6750
         TabIndex        =   16
         Top             =   300
         Width           =   510
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   2700
         TabIndex        =   12
         Text            =   "0"
         Top             =   300
         Width           =   660
      End
      Begin VB.TextBox txtPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4380
         TabIndex        =   14
         Text            =   "0.00"
         Top             =   300
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   7305
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   300
         Width           =   1290
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   11
         Top             =   300
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
         Left            =   3390
         TabIndex        =   13
         Top             =   300
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
         Caption         =   "FREE"
         Height          =   240
         Index           =   19
         Left            =   12420
         TabIndex        =   69
         Top             =   60
         Width           =   540
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Add. Charges"
         Height          =   240
         Index           =   7
         Left            =   10440
         TabIndex        =   68
         Top             =   60
         Width           =   990
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Credit Term"
         Height          =   240
         Index           =   6
         Left            =   11580
         TabIndex        =   64
         Top             =   60
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Price"
         Height          =   240
         Index           =   5
         Left            =   5610
         TabIndex        =   60
         Top             =   60
         Width           =   900
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   14
         Left            =   8670
         TabIndex        =   37
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9480
         TabIndex        =   36
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3390
         TabIndex        =   35
         Top             =   60
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
         Top             =   60
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4410
         TabIndex        =   33
         Top             =   60
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2715
         TabIndex        =   32
         Top             =   60
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   7305
         TabIndex        =   31
         Top             =   60
         Width           =   1260
      End
   End
   Begin VB.TextBox txtLocation 
      Height          =   285
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1800
      Width           =   3315
   End
   Begin VB.TextBox txtOwner 
      Height          =   285
      Left            =   1590
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2130
      Width           =   3315
   End
   Begin VB.TextBox txtRefNo 
      Height          =   285
      Left            =   7080
      TabIndex        =   6
      Top             =   1035
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   210
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   3840
      Width           =   13275
      _ExtentX        =   23416
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
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   7080
      TabIndex        =   7
      Top             =   1350
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89653251
      CurrentDate     =   38207
   End
   Begin MSDataListLib.DataCombo dcAgent 
      Height          =   315
      Left            =   1590
      TabIndex        =   1
      Top             =   1080
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdClient 
      Height          =   315
      Left            =   1590
      TabIndex        =   2
      Top             =   1440
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
      Height          =   315
      Left            =   7080
      TabIndex        =   59
      Top             =   1350
      Width           =   2505
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      Caption         =   "Additional Charges"
      Height          =   255
      Left            =   5550
      TabIndex        =   71
      Top             =   720
      Width           =   1485
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date"
      Height          =   255
      Left            =   5730
      TabIndex        =   67
      Top             =   1740
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   225
      Left            =   270
      TabIndex        =   65
      Top             =   750
      Width           =   1275
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Term"
      Height          =   225
      Left            =   270
      TabIndex        =   63
      Top             =   2490
      Width           =   1275
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Deducted"
      Height          =   255
      Left            =   5730
      TabIndex        =   62
      Top             =   2460
      Width           =   1305
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   5730
      TabIndex        =   61
      Top             =   2100
      Width           =   1305
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Client"
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
      Index           =   3
      Left            =   270
      TabIndex        =   58
      Top             =   1470
      Width           =   1275
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5100
      TabIndex        =   57
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   11730
      X2              =   13440
      Y1              =   7890
      Y2              =   7890
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
      Left            =   9960
      TabIndex        =   56
      Top             =   7530
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
      Left            =   9960
      TabIndex        =   55
      Top             =   7230
      Visible         =   0   'False
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
      Left            =   9960
      TabIndex        =   54
      Top             =   7980
      Width           =   2040
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   13500
      Y1              =   2820
      Y2              =   2820
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   13500
      Y1              =   2790
      Y2              =   2790
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
      Left            =   9960
      TabIndex        =   53
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   52
      Top             =   6600
      Width           =   990
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
      Left            =   9960
      TabIndex        =   51
      Top             =   6930
      Width           =   2040
   End
   Begin VB.Shape Shape1 
      Height          =   8295
      Left            =   120
      Top             =   570
      Width           =   13455
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8895
      Left            =   60
      Top             =   60
      Width           =   13605
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Location"
      Height          =   225
      Left            =   270
      TabIndex        =   50
      Top             =   1800
      Width           =   1275
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Booking Agent"
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
      Left            =   270
      TabIndex        =   49
      Top             =   1110
      Width           =   1275
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Issued:"
      Height          =   255
      Index           =   1
      Left            =   5730
      TabIndex        =   48
      Top             =   1380
      Width           =   1305
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner"
      Height          =   225
      Left            =   270
      TabIndex        =   47
      Top             =   2160
      Width           =   1275
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Receipts"
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
      TabIndex        =   46
      Top             =   150
      Width           =   4905
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
      Left            =   300
      TabIndex        =   45
      Top             =   3600
      Width           =   4365
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "DR#/ OR#"
      Height          =   255
      Left            =   5730
      TabIndex        =   44
      Top             =   1050
      Width           =   1305
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   120
      Top             =   120
      Width           =   13455
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   240
      Top             =   3600
      Width           =   13245
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Sales Receipts Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "Modification History"
      End
      Begin VB.Menu mnu_Return 
         Caption         =   "Return"
      End
      Begin VB.Menu mnu_Tally 
         Caption         =   "Tally Forms"
      End
      Begin VB.Menu mnu_Loading 
         Caption         =   "Loading Forms"
      End
      Begin VB.Menu mnu_Disc 
         Caption         =   "Overall Disc"
      End
      Begin VB.Menu mnu_Adjust 
         Caption         =   "Adjust"
      End
      Begin VB.Menu mnu_Prn 
         Caption         =   "Prn Ind/Prn Bat"
      End
      Begin VB.Menu mnu_Vat 
         Caption         =   "Show VAT && Taxbase"
      End
   End
End
Attribute VB_Name = "frmSalesReceiptsAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean
Public strRouteDesc         As String
Public ReceiptBatchPK       As Long

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset 'Main recordset for Invoice
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Old txtQty Value. Hold when editing qty
Dim cSalesPrice             As Currency

Private Sub btnAdd_Click(Index As Integer)
On Error GoTo err
    
    Dim RSStockUnit As New Recordset
    
    Dim intTotalOnhand          As Long
    Dim intTotalIncoming        As Integer
    Dim intTotalOnhInc          As Long  'Total of Onhand + Incoming
    Dim intExcessQty            As Integer
    
    Dim intSuggestedQty         As Integer
    Dim blnAddIncoming          As Boolean
    Dim intQtyOrdered           As Integer 'hold the value of txtQty
    Dim intCount                As Integer
    
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub
    
    If dcUnit.Text = "" Then
        MsgBox "Please select unit", vbInformation
        dcUnit.SetFocus
        Exit Sub
    End If
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 11, nsdStock.Tag)
    intStockID = nsdStock.Tag
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & intStockID & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    If toNumber(txtPrice.Text) <= 0 Then
        MsgBox "Please enter a valid sales price.", vbExclamation
        txtPrice.SetFocus
        Exit Sub
    End If
    
    intQtyOrdered = txtQty.Text

    RSStockUnit.Find "UnitID = " & dcUnit.BoundText

    If RSStockUnit!Onhand < intQtyOrdered Then GoSub GetOnhand

Continue:
    'Save to stock card
    Dim RSStockCard As New Recordset

    RSStockCard.CursorLocation = adUseClient
    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                .TextMatrix(1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(1, 2) = nsdStock.Text
                .TextMatrix(1, 3) = intQtyOrdered 'txtQty.Text
                .TextMatrix(1, 4) = dcUnit.Text
                .TextMatrix(1, 5) = toMoney(txtPrice.Text)
                .TextMatrix(1, 6) = toMoney(txtExtPrice.Text)
                .TextMatrix(1, 7) = toMoney(txtAddCharges.Text)
                .TextMatrix(1, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(1, 9) = toNumber(txtDisc.Text)
                .TextMatrix(1, 10) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(1, 11) = intStockID
                .TextMatrix(1, 12) = False
                .TextMatrix(1, 13) = txtCreditTerm2.Text
                .TextMatrix(1, 14) = changeYNValue(ckFree.Value)
            Else
AddIncoming:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(.Rows - 1, 2) = nsdStock.Text
                .TextMatrix(.Rows - 1, 3) = intQtyOrdered 'txtQty.Text
                .TextMatrix(.Rows - 1, 4) = dcUnit.Text
                .TextMatrix(.Rows - 1, 5) = toMoney(txtPrice.Text)
                .TextMatrix(.Rows - 1, 6) = toMoney(txtExtPrice.Text)
                .TextMatrix(.Rows - 1, 7) = toMoney(txtAddCharges.Text)
                .TextMatrix(.Rows - 1, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(.Rows - 1, 9) = toNumber(txtDisc.Text)
                .TextMatrix(.Rows - 1, 10) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(.Rows - 1, 11) = intStockID
                .TextMatrix(.Rows - 1, 12) = IIf(blnAddIncoming = True And intCount = 2, True, False)
                .TextMatrix(.Rows - 1, 13) = txtCreditTerm2.Text
                .TextMatrix(.Rows - 1, 14) = changeYNValue(ckFree.Value)
                
                .FillStyle = 1

                .Row = .Rows - 1
                .ColSel = 12
                If blnAddIncoming = True And intCount = 2 Then
                    .CellForeColor = vbBlue
                    
                    blnAddIncoming = False
                End If
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If .TextMatrix(CurrRow, 4) <> dcUnit.Text Then GoTo AddIncoming
                
            If MsgBox("Item already exist. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow

                'Restore back the invoice amount and discount
                cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 8))
                txtGross(2).Text = Format$(cIGross, "#,##0.00")
                cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 10))
                txtNet.Text = Format$(cIAmount, "#,##0.00")
                'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
                cDAmount = cDAmount - toNumber(toNumber(.TextMatrix(.Rows - 1, 9)) / 100) * _
                        (toNumber(toNumber(Grid.TextMatrix(.RowSel, 3)) * _
                        cSalesPrice))
                txtDesc.Text = Format$(cDAmount, "#,##0.00")

                .TextMatrix(CurrRow, 1) = nsdStock.getSelValueAt(1)
                .TextMatrix(CurrRow, 2) = nsdStock.Text
                .TextMatrix(CurrRow, 3) = intQtyOrdered 'txtQty.Text
                .TextMatrix(CurrRow, 4) = dcUnit.Text
                .TextMatrix(CurrRow, 5) = toMoney(txtPrice.Text)
                .TextMatrix(CurrRow, 6) = toMoney(txtExtPrice.Text)
                .TextMatrix(CurrRow, 7) = toMoney(txtAddCharges.Text)
                .TextMatrix(CurrRow, 8) = toMoney(txtGross(1).Text)
                .TextMatrix(CurrRow, 9) = toNumber(txtDisc.Text)
                .TextMatrix(CurrRow, 10) = toMoney(toNumber(txtNetAmount.Text))
                .TextMatrix(CurrRow, 13) = txtCreditTerm2.Text
                .TextMatrix(CurrRow, 14) = changeYNValue(ckFree.Value)

                'deduct qty from Stock Unit's table
                RSStockUnit.Filter = "UnitID = " & dcUnit.BoundText  'getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")

                RSStockUnit!Onhand = RSStockUnit!Onhand + intQtyOld

                RSStockUnit.Update
            Else
                Exit Sub
            End If
        End If
        
        RSStockCard.Filter = "StockID = " & intStockID & " AND RefNo2 = '" & txtRefNo.Text & "'"

        If RSStockCard.RecordCount = 0 Then RSStockCard.AddNew
        
        'Deduct qty solt to stock card
        RSStockCard!Type = "S"
        RSStockCard!UnitID = dcUnit.BoundText
        RSStockCard!RefNo2 = txtRefNo.Text
        RSStockCard!Pieces2 = intQtyOrdered
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        RSStockCard!SalesPrice = cSalesPrice
        RSStockCard!StockID = intStockID

        RSStockCard.Update
        
        RSStockUnit.Find "UnitID = " & dcUnit.BoundText

        'Deduct qty from highest unit breakdown if Onhand is less than qty ordered
        If RSStockUnit!Onhand < intQtyOrdered Then
            DeductOnhand intQtyOrdered, RSStockUnit!Order, True, RSStockUnit
        End If
        
        'deduct qty from Stock Unit's table
        RSStockUnit.Find "UnitID = " & dcUnit.BoundText
        
        RSStockUnit!Onhand = RSStockUnit!Onhand - intQtyOrdered
        
        RSStockUnit.Update
            
        'Add the amount to current load amount
        cIGross = cIGross + toNumber(txtGross(1).Text)
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
        'Use ExtPrice instead of Sales Price if ExtPrice is more than zero (0)
        cDAmount = cDAmount + toNumber(toNumber(txtDisc.Text) / 100) * _
                (toNumber(intQtyOrdered * _
                cSalesPrice))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        'Highlight the current row's column
        .ColSel = 13
        'Display a remove button
        If blnAddIncoming = True Then
            intQtyOrdered = intSuggestedQty
            intCount = 2
            GoSub AddIncoming
            
'            blnAddIncoming = False
        End If
        
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
    
    Exit Sub
    
GetOnhand:
    intTotalOnhInc = GetTotalQty("Total", RSStockUnit!Order, RSStockUnit!TotalQty, RSStockUnit)
    
    If intTotalOnhInc > 0 Then
    
        intTotalOnhand = GetTotalQty("Onhand", RSStockUnit!Order, RSStockUnit!Onhand, RSStockUnit)
        If intTotalOnhand >= 0 Then
        
            If intQtyOrdered > intTotalOnhand Then
                intExcessQty = intQtyOrdered - intTotalOnhand
                
                intTotalIncoming = GetTotalQty("Incoming", RSStockUnit!Order, RSStockUnit!Incoming, RSStockUnit)
                
                If intTotalIncoming > 0 And intTotalIncoming >= intExcessQty Then
                    intSuggestedQty = intExcessQty
                    With frmSuggestedQty
                        .intStockID = intStockID
                        .strProduct = nsdStock.Text
                        .intQtyOrdered = intTotalOnhand
                        .intQtySuggested = intExcessQty
                        
                        .show 1
                            
                        If .blnUseSuggestedQty = True And .blnCancel = False Then
                            blnAddIncoming = True
                            intSuggestedQty = intExcessQty
                        ElseIf .blnCancel = True Then
                            Exit Sub
                        End If
                        
                        intQtyOrdered = intTotalOnhand
                    End With
                Else
                    With frmSuggestedQty
                        .intStockID = intStockID
                        .strProduct = nsdStock.Text
                        .intQtyOrdered = intTotalOnhand
                        .intQtySuggested = intTotalIncoming
                        
                        .show 1
                            
                        If .blnUseSuggestedQty = True And .blnCancel = False Then
                            blnAddIncoming = True
                            intSuggestedQty = intTotalIncoming
                            
                            intCount = 1
                        ElseIf .blnCancel = True Then
                            Exit Sub
                        End If
                        
                        intQtyOrdered = intTotalOnhand
                    End With
                End If
            End If
        End If
    Else
        MsgBox "Insufficient qty", vbInformation
        With frmCustomersItem
            .StockID = intStockID
            
            .show 1
            RSStockUnit.Close
            
            If .blnCancel = False Then
                GoSub GetOnhand
            Else
                Exit Sub
            End If
        End With
    End If
    
    GoSub Continue
err:
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function DeductOnhand(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, RS As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    RS.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If RS!Onhand >= QtyNeededTemp Then
            If blnDeduct = False Then
                DeductOnhand = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - RS!Onhand) / RS!Qty
            
            RS.MoveFirst
            
            RS.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            RS!Onhand = RS!Onhand - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            RS.MoveFirst
            RS.Find "Order = " & OrderTemp
            
            RS!Onhand = RS!Onhand + (QtyNeededTemp * RS!Qty)
            
            RS.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                DeductOnhand = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        DeductOnhand = False
    End If
End Function

'Get the total Qty onhand, incoming and total of onhand and incoming
Private Function GetTotalQty(strField As String, Order As Integer, intOnhand As Integer, RS As Recordset)
    Dim strFieldValue As Integer
    Dim intOrder As Integer
    
    GetTotalQty = intOnhand
    
    intOrder = Order - 1
    
    Do Until intOrder < 1
        RS.MoveFirst
        RS.Find "Order = " & intOrder
        
        If strField = "Onhand" Then
            strFieldValue = RS!Onhand
        ElseIf strField = "Incoming" Then
            strFieldValue = RS!Incoming
        Else
            strFieldValue = RS!TotalQty
        End If
        
        GetTotalQty = GetTotalQty + GetTotalUnitQty(Order, intOrder, strFieldValue, RS)
        intOrder = intOrder - 1
    Loop
End Function

'This function is called by GetTotalQty Function
Private Function GetTotalUnitQty(Order As Integer, ByVal Ordertmp As Integer, intOnhand As Integer, RS As Recordset)
    GetTotalUnitQty = 1
    Do Until Order = Ordertmp
        Ordertmp = Ordertmp + 1
        
        RS.MoveNext
        
        GetTotalUnitQty = GetTotalUnitQty * RS!Qty
    Loop
    GetTotalUnitQty = intOnhand * GetTotalUnitQty
End Function

Private Function GetIncoming(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, RS As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    RS.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If RS!Incoming >= QtyNeededTemp Then
            If blnDeduct = False Then
                GetIncoming = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - RS!Incoming) / RS!Qty
            
            RS.MoveFirst
            
            RS.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            RS!Incoming = RS!Incoming - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            RS.MoveFirst
            RS.Find "Order = " & OrderTemp
            
            RS!Incoming = RS!Incoming + (QtyNeededTemp * RS!Qty)
            
            RS.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                GetIncoming = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        GetIncoming = False
    End If
End Function

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update grooss to current purchase amount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 8))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 10))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(toNumber(.TextMatrix(.Rows - 1, 9)) / 100) * (toNumber(toNumber(Grid.TextMatrix(.RowSel, 3)) * toNumber(Grid.TextMatrix(.RowSel, 5))))
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        Dim RSStockUnit As New Recordset
        
        RSStockUnit.CursorLocation = adUseClient
        RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & toNumber(Grid.TextMatrix(Grid.RowSel, 11)), CN, adOpenStatic, adLockOptimistic
        
        'deduct qty from Stock Unit's table
        RSStockUnit.Filter = "UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(Grid.RowSel, 4) & "'", "UnitID")

        RSStockUnit!Onhand = RSStockUnit!Onhand + toNumber(Grid.TextMatrix(Grid.RowSel, 3))

        RSStockUnit.Update
        
        RSStockUnit.Close
        
        'Save to stock card
        Dim RSStockCard As New Recordset

        RSStockCard.CursorLocation = adUseClient
        RSStockCard.Open "SELECT * FROM Stock_Card WHERE StockID = " & toNumber(Grid.TextMatrix(Grid.RowSel, 11)) & " AND RefNo2 = '" & txtRefNo.Text & "'", CN, adOpenStatic, adLockOptimistic
        
        RSStockCard!Pieces2 = RSStockCard!Pieces2 - toNumber(Grid.TextMatrix(Grid.RowSel, 3))
        
        RSStockCard.Update
        
        RSStockCard.Close
        
    If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cboStatus_Click()
    If cboStatus.ListIndex = 0 And ReceiptBatchPK = 0 Then 'Sold
        cmdSave.Caption = "&Payment"
    Else 'Save
        cmdSave.Caption = "&Save"
    End If
End Sub

Private Sub ckFree_Click()
    If ckFree.Value = 1 Then 'If checked
        txtDisc.Text = "0"
'        txtDisc.Visible = False
        txtGross(1).Text = "0"
'        txtGross(1).Visible = False
        txtNetAmount.Text = "0.00"
'        txtNetAmount.Visible = False
'        Labels(17).Visible = False
'        Labels(14).Visible = False
'        Label1.Visible = False
    Else
        txtQty_Change
        
        txtGross(1).Visible = True
        txtDisc.Visible = True
        txtNetAmount.Visible = True
        Labels(17).Visible = True
        Labels(14).Visible = True
        Label1.Visible = True
    End If
End Sub

Private Sub cmdPrint_Click()
    Unload frmReports

    With frmReports
        .strReport = "Receipt Form Report"
        .strWhere = "{qry_Receipt_Form.ClientID} = " & nsdClient.Tag & " AND {qry_Receipt_Form.ReceiptID} = " & PK
        
        LoadForm frmReports
    End With
End Sub

Private Sub CmdTasks_Click()
    PopupMenu mnu_Tasks
End Sub

Private Sub dcRoute_Click(Area As Integer)
    Dim strRoute As String
    
    strRoute = getValueAt("SELECT Route, RouteID FROM Routes WHERE RouteID=" & dcRoute.BoundText, "Route")
    chkAddCharges.Value = changeTFValue(CStr(getValueAt("SELECT AddCharges, RouteID FROM Routes WHERE RouteID=" & dcRoute.BoundText, "AddCharges")))
    txtRefNo.Text = strRoute & Format(Date, "yy") & Format(PK, "000000")
End Sub

Private Sub dcUnit_Change()
    If dcUnit.Text = "" Then Exit Sub
    
    txtPrice.Text = toMoney(getValueAt("SELECT SalesPrice,ExtPrice FROM qry_Stock_Unit WHERE StockID= " & nsdStock.Tag & " AND UnitID = " & dcUnit.BoundText & "", "SalesPrice"))
    cSalesPrice = txtPrice.Text
    
    txtQty_Change
'    Validate_ExtPrice
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

Private Sub mnu_Return_Click()
    Dim RSSalesReturn As New Recordset

    RSSalesReturn.CursorLocation = adUseClient
    RSSalesReturn.Open "SELECT SalesReturnID FROM Sales_Return WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    With frmSalesReturnAE
        If RSSalesReturn.RecordCount > 0 Then 'if record exist then edit record
            Dim blnStatus As Boolean
            
            blnStatus = getValueAt("SELECT SalesReturnID,Status FROM Sales_Return WHERE SalesReturnID=" & RSSalesReturn!SalesReturnID, "Status")
            
            If blnStatus Then 'true
                .State = adStateViewMode
            Else
                .State = adStateEditMode
            End If
            
            .PK = RSSalesReturn!SalesReturnID
        Else
            .State = adStateAddMode
            .ReceiptPK = PK
        End If
        
        .show vbModal
    End With
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

Private Sub nsdClient_Change()
    If nsdClient.DisableDropdown = False Then
        txtLocation.Text = nsdClient.getSelValueAt(3)
        txtOwner.Text = nsdClient.getSelValueAt(4)
        txtCreditTerm.Text = nsdClient.getSelValueAt(5)
    End If
End Sub

Private Sub txtCreditTerm2_Validate(Cancel As Boolean)
    txtCreditTerm2.Text = toNumber(txtCreditTerm2.Text)
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

Private Sub txtDisc_LostFocus()
    txtQty_Change
End Sub

Private Sub txtdisc_Validate(Cancel As Boolean)
    txtDisc.Text = toNumber(txtDisc.Text)
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    'Verify the entries
    If nsdClient.Text = "" Then
        MsgBox "Please select a client.", vbExclamation
        nsdClient.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before you can save this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
              
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Receipts_Detail WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    DeleteItems
    
    'Save the record
    With RS
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![ReceiptID] = PK
            !ReceiptBatchID = ReceiptBatchPK
            ![ClientID] = nsdClient.BoundText
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        ElseIf State = adStateEditMode Then
            .Close
            .Open "SELECT * FROM Receipts WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic
            
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        
        !RouteID = dcRoute.BoundText
        !AgentID = dcAgent.BoundText
        !RefNo = txtRefNo.Text
        !DateIssued = dtpDate.Value
        ![Status] = IIf(cboStatus.Text = "Sold", True, False)
        ![Deducted] = cboDeducted.Text
        ![Notes] = txtNotes.Text

        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)

        .Update
    End With
    
    Dim intUnitsOrder As Integer
    Dim intQty As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSDetails.AddNew

                RSDetails![ReceiptID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 11))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 4))
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![ExtPrice] = toNumber(.TextMatrix(c, 6))
                RSDetails![AddCharges] = toNumber(.TextMatrix(c, 7))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![Suggested] = .TextMatrix(c, 12)
                RSDetails![CreditTerm] = IIf(.TextMatrix(c, 13) = "", 0, .TextMatrix(c, 13))
                
                RSDetails.Update
                
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 11))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew
                
                RSDetails![ReceiptID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 11))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getUnitID(.TextMatrix(c, 4))
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![ExtPrice] = toNumber(.TextMatrix(c, 6))
                RSDetails![AddCharges] = toNumber(.TextMatrix(c, 7))
                RSDetails![Discount] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![Suggested] = .TextMatrix(c, 12)
                RSDetails![CreditTerm] = IIf(.TextMatrix(c, 13) = "", 0, .TextMatrix(c, 13))
                
                RSDetails.Update
                
            End If
            
        Next c
    End With

    If cboStatus.Text = "Sold" Then
        Dim RSClientsLedger As New Recordset
    
        RSClientsLedger.CursorLocation = adUseClient
        RSClientsLedger.Open "SELECT * FROM Clients_Ledger WHERE LedgerID=" & 0, CN, adOpenStatic, adLockOptimistic
    
        With RSClientsLedger
            .AddNew
            
            !LedgerID = getIndex("Clients_Ledger")
            !ReceiptID = PK
            !ReceiptBatchID = IIf(ReceiptBatchPK <> 0, ReceiptBatchPK, 0)
            !ClientID = IIf(nsdClient.BoundText = "", nsdClient.Tag, nsdClient.BoundText)
            !Date = dtpDeliveryDate.Value
            !RefNo = txtRefNo.Text
            !ChargeAccount = "Credit"
            !Debit = txtNet.Text
            
            .Update
        End With
    End If
    
    If cmdSave.Caption = "&Payment" And cboStatus.Text = "Sold" Then
        With frmPayment
            .PK = PK
            .ClientID = IIf(nsdClient.BoundText = "", nsdClient.Tag, nsdClient.BoundText)
            .strCustomer = nsdClient.Text
            .strRefNo = txtRefNo.Text
            .TotalAmount = txtNet.Text
            
            .show vbModal
        End With
    End If
    
    'Clear variables
    c = 0
    Set RSDetails = Nothing
'    Set RSClientsLedger = Nothing
    CN.CommitTrans
    
    blnSave = True
    
    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
               
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
            
            CN.BeginTrans
        Else
            Unload Me
        End If
    Else
        MsgBox "Changes in record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    blnSave = False
    
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        nsdClient.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    Dim strRoute As String
    
    InitGrid
    
    bind_dc "SELECT * FROM Routes", "Desc", dcRoute, "RouteID", True
    bind_dc "SELECT * FROM Agents", "AgentName", dcAgent, "AgentID", True
    
    'zero means walk-in customer
    If ReceiptBatchPK = 0 Then _
        dcRoute.BoundText = 21
        
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
    InitNSD
    
    'Set the recordset
    RS.CursorLocation = adUseClient
    If RS.State = 1 Then RS.Close
        RS.Open "SELECT * FROM Receipts WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDate.Value = Date
        mnu_Return.Enabled = False
        
        GeneratePK
        CN.BeginTrans

        If strRouteDesc <> "" Then _
            dcRoute.Text = strRouteDesc
        
        strRoute = getValueAt("SELECT Route, RouteID FROM Routes WHERE RouteID=" & dcRoute.BoundText, "Route")
        txtRefNo.Text = strRoute & Format(Date, "yy") & Format(PK, "000000")
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        RS.Open "SELECT * FROM qry_Receipts WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        cmdPrint.Visible = True
        
        If State = adStateViewMode Then
            cmdCancel.Caption = "Close"
            DisplayForViewing
        Else
            mnu_Return.Enabled = False
            InitNSD
            
            CN.BeginTrans
            
            DisplayForEditing
        End If
    
        If ForCusAcc = True Then
            Me.Icon = frmSalesReceipts.Icon
        End If

        Screen.MousePointer = vbDefault
    End If
    
    'Initialize Graphics
    With MAIN
        'cmdGenerate.Picture = .i16x16.ListImages(14).Picture
        'cmdNew.Picture = .i16x16.ListImages(10).Picture
        'cmdReset.Picture = .i16x16.ListImages(15).Picture
    End With
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Receipts")
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
        .Cols = 15
        .ColSel = 14
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 0
        .ColWidth(2) = 2505
        .ColWidth(3) = 1000
        .ColWidth(4) = 900
        .ColWidth(5) = 900
        .ColWidth(6) = 900
        .ColWidth(7) = 1200
        .ColWidth(8) = 900
        .ColWidth(9) = 1200
        .ColWidth(10) = 1200
        .ColWidth(11) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Description"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Sales Price"
        .TextMatrix(0, 6) = "Ext Price"
        .TextMatrix(0, 7) = "Add. Charges"
        .TextMatrix(0, 8) = "Gross"
        .TextMatrix(0, 9) = "Discount(%)"
        .TextMatrix(0, 10) = "Net Amount"
        .TextMatrix(0, 11) = "Stock ID"
        .TextMatrix(0, 12) = "Suggested"
        .TextMatrix(0, 13) = "Credit Term"
        .TextMatrix(0, 14) = "Free"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
'        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbLeftJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
'        .ColAlignment(9) = vbRightJustify
'        .ColAlignment(11) = vbLeftJustify
'        .ColAlignment(12) = vbRightJustify
    End With
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtPrice.Tag = 0
    txtPrice.Text = "0.00"
    txtQty.Text = 0
    txtExtPrice.Text = "0.00"
End Sub

Private Sub Form_Unload(Cancel As Integer)
'    If HaveAction = True Then
'        'frmSalesReceipts.RefreshRecords
'    End If
    
    Set frmSalesReceiptsAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If State = adStateViewMode Then Exit Sub

        nsdStock.Text = .TextMatrix(.RowSel, 2)
        nsdStock.Tag = .TextMatrix(.RowSel, 11) 'Add tag coz boundtext is empty
        intQtyOld = IIf(.TextMatrix(.RowSel, 3) = "", 0, .TextMatrix(.RowSel, 3))
        txtQty.Text = .TextMatrix(.RowSel, 3)
        
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 11), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        'disable unit to prevent user from changing it. changing of unit will result to imbalance of inventory
'        If State = adStateEditMode Then dcUnit.Enabled = False
        txtPrice.Text = toMoney(.TextMatrix(.RowSel, 5))
        txtExtPrice.Text = toMoney(.TextMatrix(.RowSel, 6))
        txtAddCharges.Text = toMoney(.TextMatrix(.RowSel, 7))
        txtGross(1).Text = toMoney(.TextMatrix(.RowSel, 8))
        txtDisc.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtNetAmount.Text = toMoney(.TextMatrix(.RowSel, 10))
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 11) = "" Then '11 = StockID
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
    
    nsdStock.Tag = nsdStock.BoundText
    txtQty.Text = "0"
    txtCreditTerm2.Text = txtCreditTerm.Text

    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText, "Unit", dcUnit, "UnitID", True
    
'    txtPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Selling Price
End Sub

Private Sub txtDate_GotFocus()
    'HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

Private Sub txtExtPrice_Change()
    cSalesPrice = toMoney(toNumber(txtExtPrice.Text))
    
    txtQty_Change
End Sub

Private Sub txtExtPrice_GotFocus()
    HLText txtExtPrice
End Sub


Private Sub txtExtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtPrice_Validate(Cancel As Boolean)
    txtExtPrice.Text = toMoney(toNumber(txtExtPrice.Text))
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtPrice_Change()
    cSalesPrice = txtPrice.Text
    
    txtQty_Change
End Sub

Private Sub txtPrice_Validate(Cancel As Boolean)
    txtPrice.Text = toMoney(toNumber(txtPrice.Text))
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnAdd(0).Enabled = False
        btnAdd(1).Enabled = False
        Exit Sub
    Else
        btnAdd(0).Enabled = True
        btnAdd(1).Enabled = True
    End If
       
    txtGross(1).Text = toMoney((toNumber(txtQty.Text) * cSalesPrice))
    txtNetAmount.Text = toMoney((toNumber(txtQty.Text) * _
            cSalesPrice) - _
            ((toNumber(txtDisc.Text) / 100) * _
            toNumber(toNumber(txtQty.Text) * _
            cSalesPrice)))
    
    Validate_ExtPrice
End Sub

Private Sub Validate_ExtPrice()
On Error Resume Next

    Dim intMinimumQty As Integer
    Dim intHighestPackaging As Integer
    
    intMinimumQty = getValueAt("SELECT Qty, StockID FROM qry_Avail_ExtPrice WHERE StockID=" & nsdStock.Tag, "Qty")
    intHighestPackaging = getValueAt("SELECT UnitID From Stock_Unit WHERE StockID=" & nsdStock.Tag & " AND [Order]=1", "UnitID")
    
    If intMinimumQty = 0 Then Exit Sub
    
    If dcUnit.BoundText <> intHighestPackaging Then Exit Sub
    
    If toNumber(txtQty.Text) >= intMinimumQty Then
        txtExtPrice.Text = nsdStock.getSelValueAt(3)
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
    
'    nsdClient.Text = ""
'    txtSONo.Text = ""
'    txtAddress.Text = ""
'    txtDate.Text = ""
'    txtSalesman.Text = ""
'    txtOrderedBy.Text = ""
'    txtDispatchedBy.Text = ""
'    txtDeliveryInstructions.Text = ""
'
'    txtGross(2).Text = "0.00"
'    txtDesc.Text = "0.00"
'    txtTaxBase.Text = "0.00"
'    txtVat.Text = "0.00"
'    txtNet.Text = "0.00"
'
'    cIAmount = 0
'    cDAmount = 0
'
'    nsdClient.SetFocus
End Sub

'Used to display record
Private Sub DisplayForEditing()
    On Error GoTo err
    
    dcRoute.BoundText = RS![RouteID]
    dcAgent.Text = RS![AgentName]
    nsdClient.Tag = RS!ClientID
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = RS![Company]
    txtLocation.Text = RS![City]
    txtOwner.Text = RS![OwnersName]
    txtCreditTerm.Text = RS![CreditTerm]
    txtRefNo.Text = RS![RefNo]
    dtpDate.Value = RS![DateIssued]
    cboStatus.Text = RS!Status_Alias
    cboDeducted.Text = RS!Deducted
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = RS![Notes]

    cIGross = toNumber(txtGross(2).Text)
    cDAmount = toNumber(txtDesc.Text)
    cIAmount = toNumber(txtNet.Text)
    cIRowCount = 0

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Receipts_Detail WHERE ReceiptID=" & PK & " ORDER BY ReceiptDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![ExtPrice])
                    .TextMatrix(1, 7) = toMoney(RSDetails![AddCharges])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![Discount] * 100
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = RSDetails![StockID]
                    .TextMatrix(1, 12) = RSDetails![Suggested]
                    .TextMatrix(1, 13) = RSDetails![CreditTerm]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![ExtPrice])
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![AddCharges])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]
                    .TextMatrix(.Rows - 1, 12) = RSDetails![Suggested]
                    .TextMatrix(.Rows - 1, 13) = RSDetails![CreditTerm]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 13
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionByRow
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    cmdSave.Caption = "Save"
    
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
    
    dcRoute.BoundText = RS![RouteID]
    dcAgent.Text = RS![AgentName]
    nsdClient.Tag = RS!ClientID
    nsdClient.DisableDropdown = True
    nsdClient.TextReadOnly = True
    nsdClient.Text = RS![Company]
    txtLocation.Text = RS![City]
    txtOwner.Text = RS![OwnersName]
    txtCreditTerm.Text = RS![CreditTerm]
    txtRefNo.Text = RS![RefNo]
    txtDate.Text = RS![DateIssued]
    cboStatus.Text = RS!Status_Alias
    cboDeducted.Text = RS!Deducted
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    txtNotes.Text = RS![Notes]

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Receipts_Detail WHERE ReceiptID=" & PK & " ORDER BY ReceiptDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![Qty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 6) = toMoney(RSDetails![ExtPrice])
                    .TextMatrix(1, 7) = toMoney(RSDetails![AddCharges])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![Discount] * 100
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = RSDetails![StockID]
                    .TextMatrix(1, 12) = RSDetails![Suggested]
                    .TextMatrix(1, 13) = RSDetails![CreditTerm]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 6) = toMoney(RSDetails![ExtPrice])
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![AddCharges])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]
                    .TextMatrix(.Rows - 1, 12) = RSDetails![Suggested]
                    .TextMatrix(.Rows - 1, 13) = RSDetails![CreditTerm]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 13
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionByRow
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
    btnAdd(0).Visible = False
    btnAdd(1).Visible = False

    'Resize and reposition the controls
    cmdPrint.Left = cmdSave.Left
    
    Shape3.Top = 2850
    Label11.Top = 2850
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 3100
    Grid.Height = 3380

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
        '.sqlwCondition = "[Routes].[Route]='" & Replace(dcRoute.Text, "'", "''") & "'"
        
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
        .AddColumn "Ext. Price", 3085.26
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,ExtPrice,StockID"
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

Private Sub txtPrice_GotFocus()
    HLText txtPrice
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSReceipts As New Recordset

    If State = adStateAddMode Then Exit Sub

    RSReceipts.CursorLocation = adUseClient
    RSReceipts.Open "SELECT * FROM Receipts_Detail WHERE ReceiptID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSReceipts.RecordCount > 0 Then
        RSReceipts.MoveFirst
        While Not RSReceipts.EOF
            CurrRow = getFlexPos(Grid, 11, RSReceipts!StockID)

            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Receipts_Detail", "ReceiptDetailID", "", True, RSReceipts!ReceiptDetailID
                End If
            End With
            RSReceipts.MoveNext
        Wend
    End If
End Sub
