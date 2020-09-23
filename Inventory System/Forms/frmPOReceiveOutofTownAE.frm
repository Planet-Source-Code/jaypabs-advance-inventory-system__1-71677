VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmForwardersReceiveAE 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   8730
   ScaleWidth      =   15225
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   210
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   77
      Tag             =   "Remarks"
      Top             =   6570
      Width           =   5910
   End
   Begin VB.TextBox txtArrivalDate 
      Height          =   315
      Left            =   6060
      TabIndex        =   54
      Top             =   2310
      Width           =   1785
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   210
      ScaleHeight     =   630
      ScaleWidth      =   14775
      TabIndex        =   23
      Top             =   3060
      Width           =   14775
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   30
         TabIndex        =   35
         Top             =   255
         Width           =   2715
      End
      Begin VB.TextBox txtExtDiscPerc 
         Height          =   315
         Left            =   10170
         TabIndex        =   34
         Text            =   "0"
         Top             =   255
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         BackColor       =   &H00E6FFFF&
         Height          =   315
         Left            =   12000
         Locked          =   -1  'True
         TabIndex        =   33
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   255
         Width           =   1035
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   13890
         TabIndex        =   32
         Top             =   270
         Width           =   840
      End
      Begin VB.TextBox txtQty 
         Height          =   315
         Left            =   2775
         TabIndex        =   31
         Text            =   "0"
         Top             =   255
         Width           =   660
      End
      Begin VB.TextBox txtPrice 
         Height          =   315
         Left            =   6810
         TabIndex        =   30
         Text            =   "0.00"
         Top             =   255
         Width           =   1185
      End
      Begin VB.TextBox txtGross 
         BackColor       =   &H00E6FFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   1
         Left            =   8055
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Text            =   "0.00"
         Top             =   255
         Width           =   1290
      End
      Begin VB.TextBox txtExtDiscAmt 
         Height          =   315
         Left            =   10950
         TabIndex        =   28
         Text            =   "0"
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtDiscPercent 
         Height          =   315
         Left            =   9390
         TabIndex        =   27
         Text            =   "0"
         Top             =   255
         Width           =   735
      End
      Begin VB.TextBox txtFreight 
         Height          =   315
         Left            =   13080
         TabIndex        =   26
         Top             =   255
         Width           =   735
      End
      Begin VB.TextBox txtLooseCargo 
         Height          =   315
         Left            =   4440
         TabIndex        =   25
         Top             =   255
         Width           =   1155
      End
      Begin VB.TextBox txtLocalArrastre 
         Height          =   315
         Left            =   5640
         TabIndex        =   24
         Top             =   255
         Width           =   1155
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   3480
         TabIndex        =   36
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
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.%"
         Height          =   240
         Index           =   14
         Left            =   10140
         TabIndex        =   48
         Top             =   60
         Width           =   840
      End
      Begin VB.Label Label1 
         Caption         =   "Net Amount"
         Height          =   255
         Left            =   12180
         TabIndex        =   47
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   3480
         TabIndex        =   46
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
         TabIndex        =   45
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   6840
         TabIndex        =   44
         Top             =   30
         Width           =   1290
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   2775
         TabIndex        =   43
         Top             =   30
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   17
         Left            =   8115
         TabIndex        =   42
         Top             =   30
         Width           =   1260
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc.Amt"
         Height          =   240
         Index           =   3
         Left            =   11010
         TabIndex        =   41
         Top             =   60
         Width           =   1020
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc.%"
         Height          =   240
         Index           =   5
         Left            =   9390
         TabIndex        =   40
         Top             =   30
         Width           =   840
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Loose Cargo"
         Height          =   240
         Index           =   4
         Left            =   4440
         TabIndex        =   39
         Top             =   0
         Width           =   1050
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight %"
         Height          =   210
         Index           =   7
         Left            =   13110
         TabIndex        =   38
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Local Arrastre"
         Height          =   240
         Index           =   1
         Left            =   5640
         TabIndex        =   37
         Top             =   0
         Width           =   1050
      End
   End
   Begin VB.CommandButton CmdReceive 
      Caption         =   "Receive Items"
      Height          =   315
      Left            =   10890
      TabIndex        =   22
      Top             =   8160
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   21
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7650
      Width           =   1425
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   225
      TabIndex        =   20
      Top             =   8160
      Width           =   1755
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   13650
      TabIndex        =   19
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   12270
      TabIndex        =   18
      Top             =   8160
      Width           =   1335
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmPOReceiveOutofTownAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   17
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
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   16
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6300
      Width           =   1425
   End
   Begin VB.CommandButton cmdPH 
      Caption         =   "Payment History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2100
      TabIndex        =   15
      Top             =   8160
      Width           =   1590
   End
   Begin VB.TextBox txtVat 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   13560
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   7200
      Width           =   1425
   End
   Begin VB.TextBox txtTaxBase 
      BackColor       =   &H00E6FFFF&
      Height          =   285
      Left            =   13560
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   6900
      Width           =   1425
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   13560
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6600
      Width           =   1425
   End
   Begin VB.TextBox txtTruckNo 
      Height          =   314
      Left            =   11790
      TabIndex        =   11
      Top             =   2340
      Width           =   1215
   End
   Begin VB.TextBox txtShippingGuideNo 
      Height          =   314
      Left            =   1860
      TabIndex        =   10
      Top             =   1590
      Width           =   1545
   End
   Begin VB.TextBox txtShip 
      Height          =   314
      Left            =   9660
      TabIndex        =   9
      Top             =   1620
      Width           =   3345
   End
   Begin VB.TextBox txtVanNo 
      Height          =   314
      Left            =   11790
      TabIndex        =   8
      Top             =   1980
      Width           =   1215
   End
   Begin VB.TextBox txtVoyageNo 
      Height          =   314
      Left            =   9660
      TabIndex        =   7
      Top             =   1980
      Width           =   1095
   End
   Begin VB.TextBox txtPONo 
      Height          =   285
      Left            =   1830
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   660
      Width           =   1905
   End
   Begin VB.TextBox txtBLNo 
      Height          =   314
      Left            =   9660
      TabIndex        =   5
      Top             =   2340
      Width           =   1095
   End
   Begin VB.TextBox txtDRNo 
      Height          =   314
      Left            =   6060
      TabIndex        =   4
      Top             =   1950
      Width           =   1785
   End
   Begin VB.TextBox txtSupplier 
      Height          =   314
      Left            =   1830
      TabIndex        =   3
      Top             =   990
      Width           =   3075
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmPOReceiveOutofTownAE.frx":01B2
      Left            =   10140
      List            =   "frmPOReceiveOutofTownAE.frx":01BC
      TabIndex        =   2
      Text            =   " "
      Top             =   630
      Width           =   2325
   End
   Begin VB.ComboBox cboClass 
      Height          =   315
      ItemData        =   "frmPOReceiveOutofTownAE.frx":01D3
      Left            =   1860
      List            =   "frmPOReceiveOutofTownAE.frx":01DD
      TabIndex        =   1
      Top             =   2340
      Width           =   3225
   End
   Begin VB.CommandButton CmdSubdivide 
      Caption         =   "Subdivide Freight Percentage Equally"
      Height          =   315
      Left            =   3780
      TabIndex        =   0
      Top             =   8160
      Width           =   3405
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   210
      TabIndex        =   49
      Top             =   8010
      Width           =   14745
      _ExtentX        =   26009
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2490
      Left            =   210
      TabIndex        =   50
      Top             =   3720
      Width           =   14775
      _ExtentX        =   26061
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
      TabIndex        =   51
      Top             =   1950
      Width           =   3240
      _ExtentX        =   5715
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
      Left            =   6060
      TabIndex        =   52
      Top             =   1590
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
   Begin MSComCtl2.DTPicker dtpArrivalDate 
      Height          =   315
      Left            =   6060
      TabIndex        =   53
      Top             =   2310
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44957699
      CurrentDate     =   38989
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   6
      Left            =   225
      TabIndex        =   78
      Top             =   6300
      Width           =   990
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
      TabIndex        =   76
      Top             =   2850
      Width           =   4365
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5130
      TabIndex        =   75
      Top             =   4290
      Width           =   1245
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   13230
      X2              =   14940
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
      Left            =   11460
      TabIndex        =   74
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
      Left            =   11460
      TabIndex        =   73
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
      Left            =   11460
      TabIndex        =   72
      Top             =   7230
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
      TabIndex        =   71
      Top             =   6930
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
      Left            =   11460
      TabIndex        =   70
      Top             =   6630
      Width           =   2040
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Local*"
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
      Left            =   5190
      TabIndex        =   69
      Top             =   1620
      Width           =   825
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      Caption         =   "Truck No."
      Height          =   255
      Left            =   10890
      TabIndex        =   68
      Top             =   2370
      Width           =   855
   End
   Begin VB.Label Label46 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Guide No.*"
      Height          =   255
      Left            =   270
      TabIndex        =   67
      Top             =   1620
      Width           =   1560
   End
   Begin VB.Label Label45 
      Alignment       =   1  'Right Justify
      Caption         =   "Class*"
      Height          =   255
      Left            =   270
      TabIndex        =   66
      Top             =   2340
      Width           =   1560
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      Caption         =   "Van No."
      Height          =   255
      Left            =   10890
      TabIndex        =   65
      Top             =   2010
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Voyage No."
      Height          =   255
      Left            =   8580
      TabIndex        =   64
      Top             =   2010
      Width           =   1035
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
      Left            =   270
      TabIndex        =   63
      Top             =   150
      Width           =   4905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Company*"
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
      TabIndex        =   62
      Top             =   1980
      Width           =   1560
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "PO No."
      Height          =   225
      Left            =   540
      TabIndex        =   61
      Top             =   690
      Width           =   1245
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   210
      X2              =   15000
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship"
      Height          =   255
      Left            =   8580
      TabIndex        =   60
      Top             =   1650
      Width           =   1035
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "B.L. No."
      Height          =   255
      Left            =   8580
      TabIndex        =   59
      Top             =   2370
      Width           =   1035
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Date*"
      Height          =   225
      Left            =   5190
      TabIndex        =   58
      Top             =   2340
      Width           =   825
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "D.R. No.*"
      Height          =   225
      Left            =   5190
      TabIndex        =   57
      Top             =   1980
      Width           =   825
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier"
      Height          =   225
      Left            =   540
      TabIndex        =   56
      Top             =   1050
      Width           =   1245
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
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   210
      X2              =   15000
      Y1              =   2760
      Y2              =   2760
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   8790
      TabIndex        =   55
      Top             =   660
      Width           =   1305
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H80000006&
      BorderWidth     =   2
      Height          =   8625
      Left            =   60
      Top             =   60
      Width           =   15135
   End
   Begin VB.Shape Shape1 
      Height          =   7995
      Left            =   150
      Top             =   600
      Width           =   14955
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   120
      Width           =   14955
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
Public ForwarderPK          As Long 'Variable used to get what record is going to edit (Invoice)
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIGross                 As Currency 'Gross Amount
Dim cIAmount                As Currency 'Current Invoice Amount
Dim cDAmount                As Currency 'Current Invoice Discount Amount
Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim intQtyOld               As Integer 'Allowed value for receive qty
Dim cLocalTrucking          As Currency
Dim cSidewalkHandling       As Currency

Private Sub btnUpdate_Click()
    Dim CurrRow As Integer
    Dim curDiscPerc As Currency
    Dim curExtDiscPerc As Currency

    If cboClass.Text = "" Or nsdLocal.Text = "" Then
        MsgBox "Class & Local Forwarder Fields needs input", vbInformation
        Exit Sub
    End If
    
    CurrRow = getFlexPos(Grid, 15, txtStock.Tag)

    'Add to grid
    With Grid
        .Row = CurrRow
        
        'Restore back the invoice amount and discount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 8))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 12))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        
        'Compute discount
        curDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 7) / 100
        curExtDiscPerc = .TextMatrix(1, 6) * .TextMatrix(1, 8) / 100
        
        cDAmount = cDAmount - (curDiscPerc + curExtDiscPerc + txtExtDiscAmt.Text)
        
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        
        .TextMatrix(CurrRow, 3) = txtQty.Text
        .TextMatrix(CurrRow, 4) = dcUnit.Text
        .TextMatrix(CurrRow, 5) = txtLooseCargo.Text
        .TextMatrix(CurrRow, 6) = txtLocalArrastre.Text
        .TextMatrix(CurrRow, 7) = toMoney(txtPrice.Text)
        .TextMatrix(CurrRow, 8) = toMoney(txtGross(1).Text)
        .TextMatrix(CurrRow, 9) = toMoney(txtDiscPercent.Text)
        .TextMatrix(CurrRow, 10) = toNumber(txtExtDiscPerc.Text)
        .TextMatrix(CurrRow, 11) = toMoney(txtExtDiscAmt.Text)
        .TextMatrix(CurrRow, 12) = toMoney(toNumber(txtNetAmount.Text))
        .TextMatrix(CurrRow, 13) = txtFreight.Text
        .TextMatrix(CurrRow, 14) = toMoney(ComputeCostPerPackage(txtQty.Text, _
                                    txtLooseCargo.Text, _
                                    txtLocalArrastre.Text, _
                                    txtFreight.Text, _
                                    CurrRow))
                
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
        .ColSel = 15
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

Private Sub cboClass_Click()
    If cboClass.Text = "Van/Truck" Then
        txtLooseCargo.Locked = True
        txtLooseCargo.BackColor = &HE6FFFF
        txtFreight.Enabled = True
        txtFreight.BackColor = &H80000005
        txtAmount(5).Locked = False
        txtAmount(5).BackColor = &H80000005
        txtAmount(6).Locked = False
        txtAmount(6).BackColor = &H80000005
        txtAmount(7).Locked = False
        txtAmount(7).BackColor = &H80000005
    Else
        txtLooseCargo.Enabled = True
        txtLooseCargo.BackColor = &H80000005
        txtFreight.Enabled = False
        txtFreight.BackColor = &HE6FFFF
        txtAmount(5).Locked = True
        txtAmount(5).BackColor = &HE6FFFF
        txtAmount(6).Locked = True
        txtAmount(6).BackColor = &HE6FFFF
        txtAmount(7).Locked = True
        txtAmount(7).BackColor = &HE6FFFF
    End If
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewerPH.INV_PK = PK
    'frmInvoiceViewerPH.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.lblTitle.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.show vbModal
End Sub

Private Sub CmdReceive_Click()
    Dim RSDetails As New Recordset
    
    RSDetails.CursorLocation = adUseClient
    
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    
    If RSDetails.RecordCount > 0 Then
        With frmPOReceiveOutofTownAE
            .State = adStateAddMode
            .PK = PK
            .show vbModal
        End With
    Else
        MsgBox "All items are already forwarded.", vbInformation
    End If
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    If txtShippingGuideNo.Text = "" Or nsdShippingCo.Text = "" Or _
        cboClass.Text = "" Or nsdLocal.Text = "" Or txtDRNo.Text = "" Then
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

    On Error GoTo err

    CN.BeginTrans

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
        ![Ship] = txtShip.Text
        ![Class] = cboClass.ListIndex
        ![LocalForwarderID] = IIf(nsdLocal.BoundText = "", nsdLocal.Tag, nsdLocal.BoundText)
        ![ArrivalDate] = dtpArrivalDate.Value
        ![DRNo] = txtDRNo.Text
        ![BLNo] = txtBLNo.Text
        ![TruckNo] = txtTruckNo.Text
        ![VanNo] = txtVanNo.Text
        ![VoyageNo] = txtVoyageNo.Text
        ![Status] = IIf(cboStatus.Text = "Received", True, False)
'        ![Notes] = txtNotes.Text
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
    
        ![DateModified] = Now
        ![LastUserFK] = CurrUser.USER_PK
                
        .Update
    End With
   
    'Connection for Transportation_Cost
    Dim RSTranspo As New Recordset

    RSTranspo.CursorLocation = adUseClient
    RSTranspo.Open "SELECT * FROM Transportation_Cost WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
   
    With RSTranspo
        If State = adStateAddMode Or State = adStatePopupMode Then
    
            .AddNew
            
            ![ForwarderID] = ForwarderPK
        End If

        ![MlaTruckingDate] = dtp(1).Value
        ![MlaTruckingOR] = txtOR(1).Text
        ![MlaTruckingAmount] = txtAmount(1).Text
        
        ![MlaArrastreDate] = dtp(2).Value
        ![MlaArrastreOR] = txtOR(2).Text
        ![MlaArrastreAmount] = txtAmount(2).Text
        
        ![MlaWfgFeeDate] = dtp(3).Value
        ![MlaWfgFeeOR] = txtOR(3).Text
        ![MlaWfgFeeAmount] = txtAmount(3).Text
        
        ![FreightDate] = dtp(4).Value
        ![FreightOR] = txtOR(4).Text
        ![FreightAmount] = txtAmount(4).Text
        
        ![LocalArrastreDate] = dtp(5).Value
        ![LocalArrastreOR] = txtOR(5).Text
        ![LocalArrastreAmount] = txtAmount(5).Text
        
        ![LocalTruckingDate] = dtp(6).Value
        ![LocalTruckingOR] = txtOR(6).Text
        ![LocalTruckingAmount] = txtAmount(6).Text
        
        ![SidewalkHandlingDate] = dtp(7).Value
        ![SidewalkHandlingOR] = txtOR(7).Text
        ![SidewalkHandlingAmount] = txtAmount(7).Text
                
        .Update
    End With
    
    'Connection for Forwarders_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Forwarders_Detail WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
          
    'Save to Purchase Order Details
    Dim RSPODetails As New Recordset

    'may be this is not needed since items received in forwarders guide is not added in inventory
    RSPODetails.CursorLocation = adUseClient
    RSPODetails.Open "SELECT * From Purchase_Order_Detail Where POID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    With Grid
        
        'Save the details of the records to Purchase_Order_Receive_Local_Detail
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
            
                RSDetails.AddNew
    
                RSDetails![ForwarderID] = ForwarderPK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 15))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![LooseCargo] = toNumber(.TextMatrix(c, 5))
                RSDetails![LocalArrastre] = toNumber(.TextMatrix(c, 6))
                RSDetails![Price] = toNumber(.TextMatrix(c, 7))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 10)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 11))
                RSDetails![FreightPercent] = toNumber(.TextMatrix(c, 13))
                RSDetails![CostPerPackage] = toNumber(.TextMatrix(c, 14))
    
                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 15))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
                RSDetails![unit] = getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                RSDetails![LooseCargo] = toNumber(.TextMatrix(c, 5))
                RSDetails![LocalArrastre] = toNumber(.TextMatrix(c, 6))
                RSDetails![Price] = toNumber(.TextMatrix(c, 7))
                RSDetails![DiscPercent] = toNumber(.TextMatrix(c, 9)) / 100
                RSDetails![ExtDiscPercent] = toNumber(.TextMatrix(c, 10)) / 100
                RSDetails![ExtDiscAmt] = toNumber(.TextMatrix(c, 11))
                RSDetails![FreightPercent] = toNumber(.TextMatrix(c, 13))
                RSDetails![CostPerPackage] = toNumber(.TextMatrix(c, 14))
    
                RSDetails.Update
                
            End If
            
            If cboStatus.Text = "Received" Then
                'add qty received in Purchase Order Details

                RSPODetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 15)), , adSearchForward, 1
                RSPODetails!QtyReceived = toNumber(RSPODetails!QtyReceived) + toNumber(.TextMatrix(c, 3))
                
                RSPODetails.Update
            End If
        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans

    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Or State = adStateEditMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdSubdivide_Click()
    Dim TotalRows As Integer
    Dim CurrRow As Integer
    
    'Count total number of rows in grid
    TotalRows = Grid.Rows - 1
    
    Do While CurrRow < TotalRows
        With Grid
        'Increase the record count
        CurrRow = CurrRow + 1
        
        .TextMatrix(CurrRow, 13) = 100 / TotalRows
        End With
    Loop
    ReCalcCPP
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
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        'Set the recordset
        rs.Open "SELECT * FROM qry_Forwarders WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpArrivalDate.Value = Date
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
                   
        DisplayForAdding
    ElseIf State = adStateEditMode Then
        'Set the recordset
        rs.Open "SELECT * FROM qry_ WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
        
        dtpArrivalDate.Value = Date
        Caption = "Edit Entry"
        cmdUsrHistory.Enabled = False
                   
        DisplayForEditing
    Else
        'Set the recordset
        rs.Open "SELECT * FROM qry_Forwarders WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
        
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
        txtStock.Tag = .TextMatrix(.RowSel, 15) 'Create tag to get the StockID
        txtQty = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtLooseCargo.Text = .TextMatrix(.RowSel, 5)
        txtLocalArrastre.Text = .TextMatrix(.RowSel, 6)
        txtPrice = toMoney(.TextMatrix(.RowSel, 7))
        txtGross(1) = toMoney(.TextMatrix(.RowSel, 8))
        txtDiscPercent.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtExtDiscPerc.Text = toMoney(.TextMatrix(.RowSel, 10))
        txtExtDiscAmt.Text = toMoney(.TextMatrix(.RowSel, 11))
        txtNetAmount = toMoney(.TextMatrix(.RowSel, 12))
        txtFreight = .TextMatrix(.RowSel, 13)
        
        If State = adStateViewMode Then Exit Sub
        If Grid.Rows = 2 And Grid.TextMatrix(1, 15) = "" Then
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

Private Sub nsdLocal_Change()
On Error GoTo err

    Dim intLocalForwarder As Integer
    
    If nsdLocal.BoundText = "" Then
        intLocalForwarder = nsdLocal.Tag
    Else
        intLocalForwarder = nsdLocal.BoundText
    End If
    
    cLocalTrucking = getValueAt("SELECT Amount FROM qry_Local_Forwarder WHERE LocalForwarderID = " & intLocalForwarder & " AND LocalForwarderAccTitleID = " & 1, "Amount")
    txtAmount(6).Text = toMoney(cLocalTrucking)
    cSidewalkHandling = getValueAt("SELECT Amount FROM qry_Local_Forwarder WHERE LocalForwarderID = " & intLocalForwarder & " AND LocalForwarderAccTitleID = " & 2, "Amount")
    txtAmount(7).Text = toMoney(cSidewalkHandling)

    Exit Sub
err:
    prompt_err err, Name, "nsdLocal_Change"
    Screen.MousePointer = vbDefault
End Sub

Private Sub nsdShippingCo_Change()
'    bind_dc "SELECT * FROM qry_Cargo_Class WHERE ShippingCompanyID=" & toNumber(nsdShippingCo.BoundText), "Cargo", dcClass, "CargoID", False 'False so that it will give an empty value
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

Private Sub txtFreight_GotFocus()
    HLText txtFreight
End Sub

Private Sub txtFreight_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtLocalArrastre_GotFocus()
    HLText txtLocalArrastre
End Sub

Private Sub txtLocalArrastre_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtLocalArrastre_Validate(Cancel As Boolean)
    txtLocalArrastre.Text = toMoney(toNumber(txtLocalArrastre.Text))
End Sub

Private Sub txtLooseCargo_GotFocus()
    HLText txtLooseCargo
End Sub

Private Sub txtLooseCargo_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtLooseCargo_Validate(Cancel As Boolean)
    txtLooseCargo.Text = toMoney(toNumber(txtLooseCargo.Text))
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtQty_LostFocus()
    Dim intQtyDue As Integer
      
    intQtyDue = getValueAt("SELECT QtyDue FROM qry_Purchase_Order_Detail WHERE POID=" & PK, "QtyDue")
    If txtQty.Text > intQtyDue Then
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
    
    intQtyOld = txtQty.Text
End Sub

Private Sub txtPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Used to edit record
Private Sub DisplayForAdding()
    On Error GoTo err
    txtSupplier.Text = rs!Company
    txtPONo.Text = rs!PONo
    
    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    
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
                If .Rows = 2 And .TextMatrix(1, 15) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![QtyDue]
                    .TextMatrix(1, 4) = RSDetails![unit]
                    .TextMatrix(1, 5) = 0
                    .TextMatrix(1, 6) = 0
                    .TextMatrix(1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 13) = 0
'                    .TextMatrix(1, 14) = RSDetails![CostperPackage]
                    .TextMatrix(1, 15) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![unit]
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = 0
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 13) = 0
'                    .TextMatrix(.Rows - 1, 14) = RSDetails![CostperPackage]
                    .TextMatrix(.Rows - 1, 15) = RSDetails![StockID]
                End If
                cIRowCount = cIRowCount + 1
            End With
            RSDetails.MoveNext
        Wend
        
        Grid.Row = 1
        Grid.ColSel = 14
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    dtpArrivalDate.Visible = True
'    txtDRDate.Visible = False
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
    txtSupplier.Text = rs!Company
    txtPONo.Text = rs!PONo
    PK = rs!POID 'get POID to be make a reference for QtyDue, etc.
    
    nsdShippingCo.Tag = rs![ShippingCompanyID]
    nsdShippingCo.Text = rs![ShippingCompany]
    txtShippingGuideNo.Text = rs![ShippingGuideNo]
    txtShip.Text = rs![Ship]
    cboClass.ListIndex = rs![Class]
    nsdLocal.Tag = rs![LocalForwarderID]
    nsdLocal.Text = rs![LocalForwarder]
    txtArrivalDate.Text = rs![ArrivalDate]
    txtDRNo.Text = rs![DRNo]
    txtBLNo.Text = rs![BLNo]
    txtTruckNo.Text = rs![TruckNo]
    txtVanNo.Text = rs![VanNo]
    txtVoyageNo.Text = rs![VoyageNo]
    cboStatus.Text = rs!Status_Alias

    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    cIRowCount = 0
    
    'Connection for Transportation_Cost
    Dim RSTranspo As New Recordset

    RSTranspo.CursorLocation = adUseClient
    RSTranspo.Open "SELECT * FROM Transportation_Cost WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
   
    With RSTranspo
        rs.Filter = "ForwarderID=" & ForwarderPK

        dtp(1).Value = ![MlaTruckingDate]
        txtOR(1).Text = ![MlaTruckingOR]
        txtAmount(1).Text = toMoney(![MlaTruckingAmount])
        
        dtp(2).Value = ![MlaArrastreDate]
        txtOR(2).Text = ![MlaArrastreOR]
        txtAmount(2).Text = toMoney(![MlaArrastreAmount])
        
        dtp(3).Value = ![MlaWfgFeeDate]
        txtOR(3).Text = ![MlaWfgFeeOR]
        txtAmount(3).Text = toMoney(![MlaWfgFeeAmount])
        
        dtp(4).Value = ![FreightDate]
        txtOR(4).Text = ![FreightOR]
        txtAmount(4).Text = toMoney(![FreightAmount])
        
        dtp(5).Value = ![LocalArrastreDate]
        txtOR(5).Text = ![LocalArrastreOR]
        txtAmount(5).Text = toMoney(![LocalArrastreAmount])
        
        dtp(6).Value = ![LocalTruckingDate]
        txtOR(6).Text = ![LocalTruckingOR]
        txtAmount(6).Text = toMoney(![LocalTruckingAmount])
        
        dtp(7).Value = ![SidewalkHandlingDate]
        txtOR(7).Text = ![SidewalkHandlingOR]
        txtAmount(7).Text = toMoney(![SidewalkHandlingAmount])
                
    End With
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & ForwarderPK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 15) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![QtyDue]
                    .TextMatrix(1, 4) = RSDetails![unit]
                    .TextMatrix(1, 5) = RSDetails![LooseCargo]
                    .TextMatrix(1, 6) = RSDetails![LocalArrastre]
                    .TextMatrix(1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 13) = RSDetails![FreightPercent]
                    .TextMatrix(1, 14) = RSDetails![CostPerPackage]
                    .TextMatrix(1, 15) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![LooseCargo]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![LocalArrastre]
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 13) = RSDetails![FreightPercent]
                    .TextMatrix(.Rows - 1, 14) = RSDetails![CostPerPackage]
                    .TextMatrix(.Rows - 1, 15) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 14
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
    txtSupplier.Text = rs!Company
    txtPONo.Text = rs!PONo
    PK = rs!POID 'get POID to be make a reference for QtyDue, etc.
    
    nsdShippingCo.Tag = rs![ShippingCompanyID]
    nsdShippingCo.Text = rs![ShippingCompany]
    txtShippingGuideNo.Text = rs![ShippingGuideNo]
    txtShip.Text = rs![Ship]
    cboClass.ListIndex = rs![Class]
    nsdLocal.Tag = rs![LocalForwarderID]
    nsdLocal.Text = rs![LocalForwarder]
    txtArrivalDate.Text = rs![ArrivalDate]
    txtDRNo.Text = rs![DRNo]
    txtBLNo.Text = rs![BLNo]
    txtTruckNo.Text = rs![TruckNo]
    txtVanNo.Text = rs![VanNo]
    txtVoyageNo.Text = rs![VoyageNo]
    cboStatus.Text = rs!Status_Alias

    txtGross(2).Text = toMoney(toNumber(rs![Gross]))
    txtDesc.Text = toMoney(toNumber(rs![Discount]))
    txtTaxBase.Text = toMoney(rs![TaxBase])
    txtVat.Text = toMoney(rs![Vat])
    txtNet.Text = toMoney(rs![NetAmount])
    
    cIGross = txtGross(2).Text
    cIAmount = txtNet.Text
    cDAmount = txtDesc.Text
    cIRowCount = 0
    
    'Connection for Transportation_Cost
    Dim RSTranspo As New Recordset

    RSTranspo.CursorLocation = adUseClient
    RSTranspo.Open "SELECT * FROM Transportation_Cost WHERE ForwarderID=" & ForwarderPK, CN, adOpenStatic, adLockOptimistic
   
    With RSTranspo
        rs.Filter = "ForwarderID=" & ForwarderPK

        dtp(1).Value = ![MlaTruckingDate]
        txtOR(1).Text = ![MlaTruckingOR]
        txtAmount(1).Text = toMoney(![MlaTruckingAmount])
        
        dtp(2).Value = ![MlaArrastreDate]
        txtOR(2).Text = ![MlaArrastreOR]
        txtAmount(2).Text = toMoney(![MlaArrastreAmount])
        
        dtp(3).Value = ![MlaWfgFeeDate]
        txtOR(3).Text = ![MlaWfgFeeOR]
        txtAmount(3).Text = toMoney(![MlaWfgFeeAmount])
        
        dtp(4).Value = ![FreightDate]
        txtOR(4).Text = ![FreightOR]
        txtAmount(4).Text = toMoney(![FreightAmount])
        
        dtp(5).Value = ![LocalArrastreDate]
        txtOR(5).Text = ![LocalArrastreOR]
        txtAmount(5).Text = toMoney(![LocalArrastreAmount])
        
        dtp(6).Value = ![LocalTruckingDate]
        txtOR(6).Text = ![LocalTruckingOR]
        txtAmount(6).Text = toMoney(![LocalTruckingAmount])
        
        dtp(7).Value = ![SidewalkHandlingDate]
        txtOR(7).Text = ![SidewalkHandlingOR]
        txtAmount(7).Text = toMoney(![SidewalkHandlingAmount])
                
    End With
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & ForwarderPK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 15) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Barcode]
                    .TextMatrix(1, 2) = RSDetails![Stock]
                    .TextMatrix(1, 3) = RSDetails![QtyDue]
                    .TextMatrix(1, 4) = RSDetails![unit]
                    .TextMatrix(1, 5) = RSDetails![LooseCargo]
                    .TextMatrix(1, 6) = RSDetails![LocalArrastre]
                    .TextMatrix(1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 13) = RSDetails![FreightPercent]
                    .TextMatrix(1, 14) = RSDetails![CostPerPackage]
                    .TextMatrix(1, 15) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Barcode]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![QtyDue]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![LooseCargo]
                    .TextMatrix(.Rows - 1, 6) = RSDetails![LocalArrastre]
                    .TextMatrix(.Rows - 1, 7) = toMoney(RSDetails![Price])
                    .TextMatrix(.Rows - 1, 8) = toMoney(RSDetails![Gross])
                    .TextMatrix(.Rows - 1, 9) = RSDetails![DiscPercent] * 100
                    .TextMatrix(.Rows - 1, 10) = RSDetails![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 11) = toMoney(RSDetails![ExtDiscAmt])
                    .TextMatrix(.Rows - 1, 12) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 13) = RSDetails![FreightPercent]
                    .TextMatrix(.Rows - 1, 14) = RSDetails![CostPerPackage]
                    .TextMatrix(.Rows - 1, 15) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 14
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

    CmdReturn.Left = cmdSave.Left
    CmdReturn.Visible = True
    
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
        .Cols = 16
        .ColSel = 15
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1560
        .ColWidth(2) = 2430
        .ColWidth(3) = 465
        .ColWidth(4) = 1000
        .ColWidth(5) = 1005
        .ColWidth(6) = 1100
        .ColWidth(7) = 690
        .ColWidth(8) = 995
        .ColWidth(9) = 1150
        .ColWidth(10) = 1000
        .ColWidth(11) = 1150
        .ColWidth(12) = 1000
        .ColWidth(13) = 600
        .ColWidth(14) = 1500
        .ColWidth(15) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Barcode"
        .TextMatrix(0, 2) = "Item"
        .TextMatrix(0, 3) = "Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Loose Cargo"
        .TextMatrix(0, 6) = "Local Arrastre"
        .TextMatrix(0, 7) = "Price" 'Supplier Price
        .TextMatrix(0, 8) = "Gross"
        .TextMatrix(0, 9) = "Disc(%)"
        .TextMatrix(0, 10) = "Ext. Disc(%)"
        .TextMatrix(0, 11) = "Ext. Disc(Amt)"
        .TextMatrix(0, 12) = "Net Amount"
        .TextMatrix(0, 13) = "Freight(%)"
        .TextMatrix(0, 14) = "Cost per Package"
        .TextMatrix(0, 15) = "Stock ID"
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
        .ColAlignment(11) = vbRightJustify
        .ColAlignment(12) = vbRightJustify
        .ColAlignment(13) = vbRightJustify
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
            CurrRow = getFlexPos(Grid, 15, RSStocks!StockID)
        
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

