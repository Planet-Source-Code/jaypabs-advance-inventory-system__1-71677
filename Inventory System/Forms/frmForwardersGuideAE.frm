VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmForwardersGuideAE 
   BorderStyle     =   0  'None
   ClientHeight    =   9345
   ClientLeft      =   -30
   ClientTop       =   -420
   ClientWidth     =   15225
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9345
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboRef 
      Height          =   315
      ItemData        =   "frmForwardersGuideAE.frx":0000
      Left            =   8850
      List            =   "frmForwardersGuideAE.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1620
      Width           =   2415
   End
   Begin VB.ComboBox cboFreightAgreement 
      Height          =   315
      ItemData        =   "frmForwardersGuideAE.frx":008F
      Left            =   6360
      List            =   "frmForwardersGuideAE.frx":00A2
      Style           =   2  'Dropdown List
      TabIndex        =   71
      Top             =   660
      Width           =   2490
   End
   Begin VB.ComboBox cboFreightPeriod 
      Height          =   315
      ItemData        =   "frmForwardersGuideAE.frx":011B
      Left            =   6360
      List            =   "frmForwardersGuideAE.frx":0125
      Style           =   2  'Dropdown List
      TabIndex        =   70
      Top             =   1035
      Width           =   2490
   End
   Begin VB.TextBox txtNotes 
      Height          =   1635
      Left            =   240
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   66
      Tag             =   "Remarks"
      Top             =   6840
      Width           =   4650
   End
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Forwarders Guide Tasks"
      Height          =   315
      Left            =   240
      TabIndex        =   65
      Top             =   8790
      Width           =   2085
   End
   Begin VB.TextBox txtPickupLocation 
      Height          =   345
      Left            =   12930
      TabIndex        =   11
      Top             =   1620
      Width           =   2115
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      ItemData        =   "frmForwardersGuideAE.frx":013C
      Left            =   1860
      List            =   "frmForwardersGuideAE.frx":0146
      TabIndex        =   5
      Top             =   2340
      Width           =   2955
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmForwardersGuideAE.frx":0162
      Left            =   10440
      List            =   "frmForwardersGuideAE.frx":016C
      TabIndex        =   2
      Text            =   " "
      Top             =   660
      Width           =   2325
   End
   Begin VB.TextBox txtSupplier 
      Height          =   314
      Left            =   1050
      TabIndex        =   1
      Top             =   990
      Width           =   3075
   End
   Begin VB.TextBox txtPONo 
      Height          =   285
      Left            =   1050
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   660
      Width           =   1905
   End
   Begin VB.TextBox txtRefNo 
      Height          =   314
      Left            =   8850
      TabIndex        =   10
      Top             =   2010
      Width           =   2415
   End
   Begin VB.TextBox txtShippingGuideNo 
      Height          =   314
      Left            =   1860
      TabIndex        =   3
      Top             =   1590
      Width           =   1545
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6930
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   13560
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   7230
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   13560
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   7530
      Visible         =   0   'False
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
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6630
      Width           =   1425
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmForwardersGuideAE.frx":0183
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Remove"
      Top             =   4050
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   12240
      TabIndex        =   22
      Top             =   8790
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   13620
      TabIndex        =   23
      Top             =   8790
      Width           =   1335
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7980
      Width           =   1425
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   840
      Left            =   210
      ScaleHeight     =   840
      ScaleWidth      =   14775
      TabIndex        =   24
      Top             =   3060
      Width           =   14775
      Begin VB.TextBox txtDiscPercent 
         Height          =   315
         Left            =   7050
         TabIndex        =   18
         Text            =   "0"
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox txtExtDiscAmt 
         Height          =   315
         Left            =   8610
         TabIndex        =   20
         Text            =   "0"
         Top             =   465
         Width           =   705
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   5715
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   465
         Width           =   1290
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   4470
         TabIndex        =   17
         Text            =   "0.00"
         Top             =   465
         Width           =   1185
      End
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   2775
         TabIndex        =   15
         Text            =   "0"
         Top             =   465
         Width           =   660
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   10470
         TabIndex        =   21
         Top             =   480
         Width           =   840
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   315
         Left            =   9360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   465
         Width           =   1035
      End
      Begin VB.TextBox txtExtDiscPerc 
         Height          =   315
         Left            =   7830
         TabIndex        =   19
         Text            =   "0"
         Top             =   465
         Width           =   735
      End
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   30
         TabIndex        =   14
         Top             =   465
         Width           =   2715
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   16
         Top             =   465
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
         Left            =   7020
         TabIndex        =   35
         Top             =   150
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.Amt"
         Height          =   360
         Index           =   3
         Left            =   8640
         TabIndex        =   34
         Top             =   60
         Width           =   690
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   5745
         TabIndex        =   33
         Top             =   150
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2760
         TabIndex        =   32
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   4470
         TabIndex        =   31
         Top             =   150
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
         Left            =   60
         TabIndex        =   30
         Top             =   180
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   29
         Top             =   210
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   9510
         TabIndex        =   28
         Top             =   180
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.%"
         Height          =   240
         Index           =   14
         Left            =   7770
         TabIndex        =   27
         Top             =   180
         Width           =   840
      End
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   270
      TabIndex        =   42
      Top             =   8670
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2580
      Left            =   210
      TabIndex        =   43
      Top             =   3960
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   4551
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
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      Enabled         =   -1  'True
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
      Left            =   6030
      TabIndex        =   6
      Top             =   1620
      Width           =   1800
      _ExtentX        =   3175
      _ExtentY        =   556
      Enabled         =   -1  'True
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
   Begin MSComCtl2.DTPicker dtpPickupDate 
      Height          =   345
      Left            =   12930
      TabIndex        =   12
      Top             =   2010
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   609
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20643843
      CurrentDate     =   38989
   End
   Begin VB.TextBox txtPickupDate 
      Height          =   345
      Left            =   12930
      TabIndex        =   13
      Top             =   2010
      Width           =   2115
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   315
      Left            =   6030
      TabIndex        =   7
      Top             =   2010
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20643843
      CurrentDate     =   38989
   End
   Begin VB.TextBox txtDeliveryDate 
      Height          =   315
      Left            =   6030
      TabIndex        =   62
      Top             =   2010
      Width           =   1785
   End
   Begin MSComCtl2.DTPicker dtpReceiptDate 
      Height          =   315
      Left            =   6030
      TabIndex        =   8
      Top             =   2370
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20643843
      CurrentDate     =   38989
   End
   Begin VB.TextBox txtReceiptDate 
      Height          =   315
      Left            =   6030
      TabIndex        =   73
      Top             =   2370
      Width           =   1785
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Receipt Date"
      Height          =   225
      Left            =   4920
      TabIndex        =   72
      Top             =   2400
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Agreement"
      Height          =   240
      Index           =   6
      Left            =   4200
      TabIndex        =   69
      Top             =   675
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Period"
      Height          =   240
      Index           =   1
      Left            =   4200
      TabIndex        =   68
      Top             =   1035
      Width           =   2115
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   240
      TabIndex        =   67
      Top             =   6570
      Width           =   990
   End
   Begin VB.Label Label49 
      Alignment       =   1  'Right Justify
      Caption         =   "Pickup Date"
      Height          =   315
      Left            =   11700
      TabIndex        =   64
      Top             =   2040
      Width           =   1185
   End
   Begin VB.Label Label48 
      Alignment       =   1  'Right Justify
      Caption         =   "Pickup Location"
      Height          =   315
      Left            =   11700
      TabIndex        =   63
      Top             =   1620
      Width           =   1185
   End
   Begin VB.Shape Shape1 
      Height          =   8625
      Left            =   150
      Top             =   600
      Width           =   14955
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   9255
      Left            =   60
      Top             =   60
      Width           =   15135
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   9090
      TabIndex        =   61
      Top             =   690
      Width           =   1305
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   15000
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   210
      X2              =   15000
      Y1              =   2790
      Y2              =   2790
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier"
      Height          =   225
      Left            =   210
      TabIndex        =   60
      Top             =   1050
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Reference"
      Height          =   285
      Left            =   7950
      TabIndex        =   59
      Top             =   1620
      Width           =   825
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date"
      Height          =   225
      Left            =   4920
      TabIndex        =   58
      Top             =   2040
      Width           =   1065
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   210
      X2              =   15000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   210
      X2              =   15000
      Y1              =   1470
      Y2              =   1470
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   210
      TabIndex        =   57
      Top             =   690
      Width           =   795
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
      TabIndex        =   56
      Top             =   1980
      Width           =   1560
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Forwarders Guide"
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
      TabIndex        =   55
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Ref No."
      Height          =   255
      Left            =   7950
      TabIndex        =   54
      Top             =   2040
      Width           =   825
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      Caption         =   "Class"
      Height          =   255
      Left            =   270
      TabIndex        =   53
      Top             =   2370
      Width           =   1560
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Guide No."
      Height          =   255
      Left            =   270
      TabIndex        =   52
      Top             =   1620
      Width           =   1560
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
      Left            =   5160
      TabIndex        =   51
      Top             =   1650
      Width           =   825
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
      Left            =   11460
      TabIndex        =   50
      Top             =   6960
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
      Left            =   11460
      TabIndex        =   49
      Top             =   7260
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
      Left            =   11460
      TabIndex        =   48
      Top             =   7560
      Visible         =   0   'False
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
      Left            =   11460
      TabIndex        =   47
      Top             =   6660
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
      Left            =   11460
      TabIndex        =   46
      Top             =   8010
      Width           =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   13230
      X2              =   14940
      Y1              =   7920
      Y2              =   7920
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5130
      TabIndex        =   45
      Top             =   4290
      Width           =   1245
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
      TabIndex        =   44
      Top             =   2850
      Width           =   4365
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   120
      Width           =   14955
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2820
      Width           =   14775
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Forwarders Guide Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "Modification History"
      End
      Begin VB.Menu mnu_ReceiveItems 
         Caption         =   "Receive Items"
      End
      Begin VB.Menu mnu_Vat 
         Caption         =   "Show VAT && Taxbase"
      End
   End
End
Attribute VB_Name = "frmForwardersGuideAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit (PO)
Public ForwarderPK          As Long 'Variable used to get what record is going to edit (Invoice)
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
Dim blnSave                 As Boolean

Private Sub btnUpdate_Click()
On Error GoTo err_btnUpdate_Click

    Dim CurrRow As Integer
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency
    Dim intQty As Integer
        
    CurrRow = getFlexPos(Grid, 11, txtStock.Tag)

    'Add to grid
    With Grid
        .Row = CurrRow
        
        'Restore back the invoice amount and discount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 6))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
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
        RSPODetails.Open "SELECT * From Purchase_Order_Detail Where POID = " & PK, CN, adOpenStatic, adLockOptimistic
        
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
    
    Exit Sub

err_btnUpdate_Click:
    MsgBox err.Number & " " & err.Description
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
        curDiscPerc = .TextMatrix(.RowSel, 8) * .TextMatrix(.RowSel, 8) / 100
        curExtDiscPerc = .TextMatrix(.RowSel, 8) * .TextMatrix(.RowSel, 10) / 100
        
        cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        Dim RSDetails As New Recordset
    
        RSDetails.CursorLocation = adUseClient
        RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND StockID = " & .TextMatrix(.RowSel, 11), CN, adOpenDynamic, adLockOptimistic
        
        If RSDetails.RecordCount > 0 Then
            'restore back qty that was added in DisplayForAdding procedure
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

Private Sub cmdSave_Click()
    On Error GoTo err
    
    'Verify the entries
    If txtShippingGuideNo.Text = "" Or nsdShippingCo.Text = "" Or _
        cboClass.Text = "" Or nsdLocal.Text = "" Or txtRefNo.Text = "" Then
        MsgBox "Please don't leave Field with asterisk (*) blank", vbInformation
        txtShippingGuideNo.SetFocus
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
    RSReceive.Open "SELECT * FROM Forwarders WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    DeleteItems
    
    'Save the record
    With RSReceive
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            
            ForwarderPK = getIndex("Forwarders")
            ![ForwarderID] = ForwarderPK
            ![POID] = PK
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        End If
        
        ![ShippingCompanyID] = IIf(nsdShippingCo.BoundText = "", nsdShippingCo.Tag, nsdShippingCo.BoundText)
        ![ShippingGuideNo] = txtShippingGuideNo.Text
'        ![Ship] = txtShip.Text
        ![Class] = cboClass.ListIndex
        ![LocalForwarderID] = IIf(nsdLocal.BoundText = "", nsdLocal.Tag, nsdLocal.BoundText)
        ![DeliveryDate] = dtpDeliveryDate.Value
        ![ReceiptDate] = dtpReceiptDate.Value
        ![Ref] = cboRef.Text
        ![RefNo] = txtRefNo.Text
'        ![TruckNo] = txtTruckNo.Text
'        ![VanNo] = txtVanNo.Text
'        ![VoyageNo] = txtVoyageNo.Text
        ![PickupLocation] = txtPickupLocation.Text
        ![PickupDate] = dtpPickupDate.Value
        ![Status] = IIf(cboStatus.Text = "Received", True, False)
        ![Notes] = txtNotes.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
    
'        ![Freight] = txtFreight.Text
'        ![Arrastre] = txtArrastre.Text
        
        ![DateModified] = Now
        ![LastUserFK] = CurrUser.USER_PK
                
        .Update
    End With

'    If cboStatus.Text = "Received" Then
'        'Connection for Vendors_Ledger
'        Dim RSLedger As New Recordset
'
'        With RSLedger
'            .CursorLocation = adUseClient
'            .Open "SELECT * FROM Vendors_Ledger WHERE ForwarderID=" & ForwarderPK & " AND BillType = 'Products'", CN, adOpenStatic, adLockOptimistic
'
'            .AddNew
'
'            !ForwarderID = ForwarderPK
'
'            !VendorID = txtSupplier.Tag
'            !RefNo = txtRefNo.Text
'            !Date = Date
'            !Debit = txtGross(2).Text
'            !BillType = "Products"
'
'            .Update
'
'            '-------------------------
'            'Save freight
'            .Close
'            .Open "SELECT * FROM Vendors_Ledger WHERE ForwarderID=" & ForwarderPK & " AND BillType = 'Freight'", CN, adOpenStatic, adLockOptimistic
'
'            'Save bill to Vendors_Ledger table
'            If cboFreightAgreement.Text = "By supplier until freight" Then
'                If cboFreightPeriod.Text = "Postpaid" Then
'                    .AddNew
'
'                    !ForwarderID = ForwarderPK
'
'                    !Credit = txtFreight.Text
'
'                    CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, ForwarderID, RefNo, [Date], Debit ) " _
'                            & "VALUES (" & nsdShippingCo.Tag & ", " & ForwarderPK & ", " & txtRefNo.Text & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
'                End If
'
'                CN.Execute "INSERT INTO Local_Forwarder_Ledger ( ForwarderID, [Date], Debit ) " _
'                        & "VALUES (" & ForwarderPK & ",#" & Date & "#, " & txtArrastre.Text & ")"
'            ElseIf cboFreightAgreement.Text = "By supplier until local arrastre" Then
'                If cboFreightPeriod.Text = "Postpaid" Then
'                    .AddNew
'
'                    !ForwarderID = ForwarderPK
'
'                    !Credit = toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)
'
'                    CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, ForwarderID, RefNo, [Date], Debit ) " _
'                            & "VALUES (" & nsdShippingCo.Tag & ", " & ForwarderPK & ", " & txtRefNo.Text & ",#" & Date & "#, " & toNumber(txtFreight.Text) & ")"
'                End If
'            ElseIf cboFreightAgreement.Text = "Half until freight" Then
'                .AddNew
'
'                !ForwarderID = ForwarderPK
'
'                !VendorID = txtSupplier.Tag
'                !Date = Date
'                !BillType = "Freight"
'
'                If cboFreightPeriod.Text = "Prepaid" Then
'                    !Dedit = toMoney(txtFreight.Text) / 2
'                Else 'Postpaid
'                    !Credit = toMoney(txtFreight.Text) / 2
'
'                    CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, ForwarderID, RefNo, [Date], Debit ) " _
'                            & "VALUES (" & nsdShippingCo.Tag & ", " & ForwarderPK & ", " & txtRefNo.Text & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
'                End If
'
'                CN.Execute "INSERT INTO Local_Forwarder_Ledger ( ForwarderID, [Date], Debit ) " _
'                        & "VALUES (" & ForwarderPK & ",#" & Date & "#, " & txtArrastre.Text & ")"
'
'                .Update
'            ElseIf cboFreightAgreement.Text = "Half until local arrastre" Then
'                .AddNew
'
'                !ForwarderID = ForwarderPK
'
'                !VendorID = txtSupplier.Tag
'                !Date = Date
'                !BillType = "Freight"
'
'                If cboFreightPeriod.Text = "Prepaid" Then
'                    !Dedit = (toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)) / 2
'                Else 'Postpaid
'                    !Credit = (toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)) / 2
'
'                    CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, ForwarderID, RefNo, [Date], Debit ) " _
'                            & "VALUES (" & nsdShippingCo.Tag & ", " & ForwarderPK & ", " & txtRefNo.Text & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
'
'                    CN.Execute "INSERT INTO Local_Forwarder_Ledger ( ForwarderID, [Date], Debit ) " _
'                            & "VALUES (" & ForwarderPK & ",#" & Date & "#, " & toNumber(txtArrastre.Text) & ")"
'                End If
'
'                .Update
'            ElseIf cboFreightAgreement.Text = "By VTM" Then
'                    CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, ForwarderID, RefNo, [Date], Debit ) " _
'                            & "VALUES (" & nsdShippingCo.Tag & ", " & ForwarderPK & ", " & txtRefNo.Text & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
'
'                    CN.Execute "INSERT INTO Local_Forwarder_Ledger ( ForwarderID, [Date], Debit ) " _
'                            & "VALUES (" & ForwarderPK & ",#" & Date & "#, " & toNumber(txtArrastre.Text) & ")"
'            End If
'        End With
'    End If
    
    'Connection for Forwarders_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Forwarders_Detail WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
          
    'Add qty ordered to qty onhand
    Dim RSStockUnit As New Recordset

    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * From Stock_Unit", CN, adOpenStatic, adLockOptimistic

    With Grid
        'Save the details of the records to Purchase_Order_Receive_Local_Detail
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                RSDetails.AddNew
    
                RSDetails![ForwarderID] = ForwarderPK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 11))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
    
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 11))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![Unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![Price] = toNumber(.TextMatrix(c, 5))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 7)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 8)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 9))
    
                RSDetails.Update
                
            End If
                      
            If cboStatus.Text = "Received" Then
                RSStockUnit.Filter = "StockID = " & toNumber(.TextMatrix(c, 11)) & " AND UnitID = " & getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                                  
                RSStockUnit!Pending = RSStockUnit!Pending + toNumber(.TextMatrix(c, 3))
                RSStockUnit.Update
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

    Screen.MousePointer = vbHourglass
    
    RS.CursorLocation = adUseClient
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDeliveryDate.Value = Date
        dtpReceiptDate.Value = Date
        dtpPickupDate.Value = Date
        mnu_History.Enabled = False
        mnu_ReceiveItems.Visible = False
        
        txtShippingGuideNo = GeneratePK()
        
        CN.BeginTrans
        
        DisplayForAdding
    ElseIf State = adStateEditMode Then
        'Set the recordset
        RS.Open "SELECT * FROM qry_Forwarders WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
        
'        dtpDeliveryDate.Value = Date
        
        mnu_History.Enabled = False
        mnu_ReceiveItems.Visible = False

        CN.BeginTrans

        DisplayForEditing
    Else
        'Set the recordset
        RS.Open "SELECT * FROM qry_Forwarders WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
        
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
    
    Set frmForwardersGuideAE = Nothing
End Sub

Private Sub Grid_Click()
    If State = adStateViewMode Then Exit Sub
    
    With Grid
        txtStock.Text = .TextMatrix(.RowSel, 2)
        txtStock.Tag = .TextMatrix(.RowSel, 11) 'Create tag to get the StockID
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
        If Grid.Rows = 2 And Grid.TextMatrix(1, 11) = "" Then
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

Private Sub mnu_ReceiveItems_Click()
    Dim RSDetails As New Recordset
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & ForwarderPK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    
    If RSDetails.RecordCount > 0 Then
        With frmForwardersReceiveAE
            .State = adStateAddMode
            .PK = ForwarderPK
            .show vbModal
        End With
    Else
        MsgBox "All items are already forwarded.", vbInformation
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
    
    txtSupplier.Tag = RS!VendorID
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo
    cboFreightAgreement.Text = RS!FreightAgreement
    cboFreightPeriod.Text = RS!FreightPeriod
    
    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])

    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    
'    Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenDynamic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 11) = "" Then
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
                    .TextMatrix(1, 11) = RSDetails![StockID]

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
                    .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]

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
  
    dtpDeliveryDate.Visible = True
'    txtDRDate.Visible = False
    lblStatus.Visible = False
    cboStatus.Visible = False
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Number & " " & err.Description
    End If
End Sub

'Used to edit record
Private Sub DisplayForEditing()
    On Error GoTo err
    txtSupplier.Tag = RS!VendorID
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo
    PK = RS!POID 'get POID to make a reference for QtyDue, etc.
    cboFreightAgreement.Text = RS!FreightAgreement
    cboFreightPeriod.Text = RS!FreightPeriod
    
    nsdShippingCo.Tag = RS![ShippingCompanyID]
    nsdShippingCo.Text = RS![ShippingCompany]
    txtShippingGuideNo.Text = RS![ShippingGuideNo]
'    txtShip.Text = rs![Ship]
    cboClass.ListIndex = RS![Class]
    nsdLocal.Tag = RS![LocalForwarderID]
    nsdLocal.Text = RS![LocalForwarder]
    dtpDeliveryDate.Value = RS![DeliveryDate]
    dtpReceiptDate.Value = RS![ReceiptDate]
    cboRef.Text = RS![Ref]
    txtRefNo.Text = RS![RefNo]
'    txtTruckNo.Text = rs![TruckNo]
'    txtVanNo.Text = rs![VanNo]
'    txtVoyageNo.Text = rs![VoyageNo]
    txtPickupLocation.Text = RS![PickupLocation]
    dtpPickupDate.Value = RS![PickupDate]
    cboStatus.Text = RS!Status_Alias

    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    
'    txtFreight.Text = toMoney(RS![Freight])
'    txtArrastre.Text = toMoney(RS![Arrastre])
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    cIRowCount = 0
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & ForwarderPK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
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
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = RSDetails![StockID]
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
                    .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]
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
    
    txtSupplier.Tag = RS!VendorID
    txtSupplier.Text = RS!Company
    txtPONo.Text = RS!PONo
    PK = RS!POID 'get POID to make a reference for QtyDue, etc.
    cboFreightAgreement.Text = RS!FreightAgreement
    cboFreightPeriod.Text = RS!FreightPeriod
    
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
'    txtTruckNo.Text = rs![TruckNo]
'    txtVanNo.Text = rs![VanNo]
'    txtVoyageNo.Text = rs![VoyageNo]
    txtPickupLocation.Text = RS![PickupLocation]
    txtPickupDate.Text = RS![PickupDate]
    cboStatus.Text = RS!Status_Alias

    txtGross(2).Text = toMoney(toNumber(RS![Gross]))
    txtDesc.Text = toMoney(toNumber(RS![Discount]))
    txtTaxBase.Text = toMoney(RS![TaxBase])
    txtVat.Text = toMoney(RS![Vat])
    txtNet.Text = toMoney(RS![NetAmount])
    
'    txtFreight.Text = toMoney(RS![Freight])
'    txtArrastre.Text = toMoney(RS![Arrastre])
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    cIRowCount = 0
        
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & ForwarderPK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
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
                    .TextMatrix(1, 6) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 7) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 8) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 10) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 11) = RSDetails![StockID]
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
                    .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]
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

    dtpDeliveryDate.Visible = False
    txtDeliveryDate.Visible = True
    dtpReceiptDate.Visible = False
    txtReceiptDate.Visible = True
    dtpPickupDate.Visible = False
    txtPickupDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnUpdate.Visible = False

    mnu_ReceiveItems.Visible = True
    
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

'Procedure used to generate PK
Private Function GeneratePK()
    GeneratePK = getIndex("Shipping_Guide")
End Function

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
        .ColWidth(5) = 690
        .ColWidth(6) = 995
        .ColWidth(7) = 1150
        .ColWidth(8) = 1000
        .ColWidth(9) = 1150
        .ColWidth(10) = 1000
        .ColWidth(11) = 0
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
        .TextMatrix(0, 11) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbRightJustify
        .ColAlignment(5) = vbRightJustify
        .ColAlignment(6) = vbRightJustify
        .ColAlignment(7) = vbRightJustify
        .ColAlignment(8) = vbRightJustify
        .ColAlignment(9) = vbRightJustify
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
    RSStocks.Open "SELECT * FROM Forwarders_Detail WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
    If RSStocks.RecordCount > 0 Then
        RSStocks.MoveFirst
        While Not RSStocks.EOF
            CurrRow = getFlexPos(Grid, 11, RSStocks!StockID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Forwarders_Detail", "ForwarderDetailID", "", True, RSStocks!ForwarderDetailID
                End If
            End With
            RSStocks.MoveNext
        Wend
    End If
End Sub


