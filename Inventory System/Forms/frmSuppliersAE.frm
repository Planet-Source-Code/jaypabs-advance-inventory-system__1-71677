VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSuppliersAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   8085
   ClientLeft      =   2190
   ClientTop       =   1335
   ClientWidth     =   10890
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   10890
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboFreightPeriod 
      Height          =   315
      ItemData        =   "frmSuppliersAE.frx":0000
      Left            =   2310
      List            =   "frmSuppliersAE.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1604
      Width           =   2490
   End
   Begin VB.ComboBox cboCreditOption 
      Height          =   315
      ItemData        =   "frmSuppliersAE.frx":0021
      Left            =   6360
      List            =   "frmSuppliersAE.frx":002E
      TabIndex        =   15
      Text            =   "Upon Order"
      Top             =   1980
      Width           =   2265
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   15
      Left            =   6360
      TabIndex        =   12
      Text            =   "0"
      Top             =   891
      Width           =   885
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   12
      Left            =   6360
      TabIndex        =   13
      Text            =   "0"
      Top             =   1254
      Width           =   1245
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   13
      Left            =   6360
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   1617
      Width           =   1245
   End
   Begin VB.TextBox txtEntry 
      Height          =   1005
      Index           =   14
      Left            =   5100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2640
      Width           =   5655
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   11
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2316
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2657
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   6360
      TabIndex        =   11
      Top             =   528
      Width           =   2475
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   9
      Top             =   3345
      Width           =   2490
   End
   Begin VB.ComboBox cboFreightAgreement 
      Height          =   315
      ItemData        =   "frmSuppliersAE.frx":005B
      Left            =   2310
      List            =   "frmSuppliersAE.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1233
      Width           =   2490
   End
   Begin VB.ComboBox cboDelivery 
      Height          =   315
      ItemData        =   "frmSuppliersAE.frx":00E7
      Left            =   2310
      List            =   "frmSuppliersAE.frx":00F1
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   862
      Width           =   2490
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "&Modification History"
      Height          =   315
      Left            =   180
      TabIndex        =   31
      Top             =   7650
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   8
      Top             =   2998
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   5
      Top             =   1975
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Company"
      Top             =   150
      Width           =   2520
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9285
      TabIndex        =   25
      Top             =   7680
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   7830
      TabIndex        =   24
      Top             =   7680
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   6360
      TabIndex        =   10
      Top             =   165
      Width           =   2475
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   150
      TabIndex        =   32
      Top             =   7545
      Width           =   10485
      _ExtentX        =   18494
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcLocation 
      Height          =   315
      Left            =   2310
      TabIndex        =   1
      Top             =   491
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3525
      Left            =   210
      TabIndex        =   45
      Top             =   3870
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   6218
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Products"
      TabPicture(0)   =   "frmSuppliersAE.frx":0108
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Labels(16)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Labels(15)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Labels(14)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Labels(8)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Labels(9)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Labels(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "GridStocks"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "nsdStock"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtDiscPercent"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtExtPrice"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "btnRemoveStock"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtExtDiscPercent"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtExtDiscAmount"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtUnitPrice"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAddStock"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Banks"
      TabPicture(1)   =   "frmSuppliersAE.frx":0124
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label4"
      Tab(1).Control(3)=   "Label1"
      Tab(1).Control(4)=   "nsdBank"
      Tab(1).Control(5)=   "GridBanks"
      Tab(1).Control(6)=   "txtAcctNo"
      Tab(1).Control(7)=   "txtAcctName"
      Tab(1).Control(8)=   "btnAddBank"
      Tab(1).Control(9)=   "txtBranch"
      Tab(1).Control(10)=   "btnRemoveBank"
      Tab(1).ControlCount=   11
      Begin VB.CommandButton btnRemoveBank 
         Height          =   275
         Left            =   -74790
         Picture         =   "frmSuppliersAE.frx":0140
         Style           =   1  'Graphical
         TabIndex        =   58
         ToolTipText     =   "Remove"
         Top             =   1140
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   -72540
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   690
         Width           =   1875
      End
      Begin VB.CommandButton btnAddBank 
         Caption         =   "Add"
         Height          =   315
         Left            =   -66690
         TabIndex        =   30
         Top             =   660
         Width           =   765
      End
      Begin VB.TextBox txtAcctName 
         Height          =   285
         Left            =   -68700
         TabIndex        =   29
         Top             =   690
         Width           =   1965
      End
      Begin VB.TextBox txtAcctNo 
         Height          =   285
         Left            =   -70620
         TabIndex        =   28
         Top             =   690
         Width           =   1875
      End
      Begin VB.CommandButton cmdAddStock 
         Caption         =   "Add"
         Height          =   315
         Left            =   9570
         TabIndex        =   23
         Top             =   630
         Width           =   840
      End
      Begin VB.TextBox txtUnitPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7110
         TabIndex        =   21
         Text            =   "0.00"
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox txtExtDiscAmount 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   5880
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox txtExtDiscPercent 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4800
         TabIndex        =   19
         Text            =   "0.00"
         Top             =   660
         Width           =   1035
      End
      Begin VB.CommandButton btnRemoveStock 
         Height          =   275
         Left            =   210
         Picture         =   "frmSuppliersAE.frx":02F2
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Remove"
         Top             =   1140
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.TextBox txtExtPrice 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   8340
         TabIndex        =   22
         Text            =   "0.00"
         Top             =   660
         Width           =   1185
      End
      Begin VB.TextBox txtDiscPercent 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3720
         TabIndex        =   18
         Text            =   "0.00"
         Top             =   660
         Width           =   1035
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   150
         TabIndex        =   17
         Top             =   630
         Width           =   3540
         _ExtentX        =   6244
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridStocks 
         Height          =   2310
         Left            =   150
         TabIndex        =   51
         ToolTipText     =   "Double item's to view landed cost"
         Top             =   990
         Width           =   10245
         _ExtentX        =   18071
         _ExtentY        =   4075
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridBanks 
         Height          =   2280
         Left            =   -74850
         TabIndex        =   59
         Top             =   1020
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   4022
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
      Begin ctrlNSDataCombo.NSDataCombo nsdBank 
         Height          =   315
         Left            =   -74850
         TabIndex        =   26
         Top             =   660
         Width           =   2250
         _ExtentX        =   3969
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
      Begin VB.Label Label1 
         Caption         =   "Bank"
         Height          =   225
         Left            =   -74850
         TabIndex        =   63
         Top             =   450
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Account Name"
         Height          =   255
         Left            =   -68700
         TabIndex        =   62
         Top             =   450
         Width           =   1365
      End
      Begin VB.Label Label3 
         Caption         =   "Acct. No."
         Height          =   225
         Left            =   -70620
         TabIndex        =   61
         Top             =   480
         Width           =   885
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Branch"
         Height          =   225
         Left            =   -72510
         TabIndex        =   60
         Top             =   480
         Width           =   525
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   13
         Left            =   180
         TabIndex        =   57
         Top             =   420
         Width           =   1515
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Supplier Price"
         Height          =   240
         Index           =   9
         Left            =   7140
         TabIndex        =   56
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc in Amt"
         Height          =   210
         Index           =   8
         Left            =   5880
         TabIndex        =   55
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Disc in %"
         Height          =   240
         Index           =   14
         Left            =   4680
         TabIndex        =   54
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ext. Price"
         Height          =   240
         Index           =   15
         Left            =   8370
         TabIndex        =   53
         Top             =   420
         Width           =   1140
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Disc in %"
         Height          =   240
         Index           =   16
         Left            =   3600
         TabIndex        =   52
         Top             =   420
         Width           =   1140
      End
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Period"
      Height          =   240
      Index           =   17
      Left            =   90
      TabIndex        =   64
      Top             =   1604
      Width           =   2115
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Option"
      Height          =   255
      Left            =   5130
      TabIndex        =   49
      Top             =   2010
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "General Disc (%)"
      Height          =   255
      Left            =   5130
      TabIndex        =   48
      Top             =   903
      Width           =   1185
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Limit"
      Height          =   255
      Left            =   5130
      TabIndex        =   47
      Top             =   1641
      Width           =   1185
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Term"
      Height          =   255
      Left            =   5130
      TabIndex        =   46
      Top             =   1272
      Width           =   1185
   End
   Begin VB.Label Label7 
      Caption         =   "Notes"
      Height          =   255
      Left            =   5130
      TabIndex        =   44
      Top             =   2340
      Width           =   1185
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   240
      Index           =   7
      Left            =   90
      TabIndex        =   43
      Top             =   2316
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cellphone"
      Height          =   240
      Index           =   5
      Left            =   90
      TabIndex        =   42
      Top             =   2672
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City"
      Height          =   255
      Index           =   0
      Left            =   5130
      TabIndex        =   41
      Top             =   534
      Width           =   1185
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Location"
      Height          =   240
      Index           =   11
      Left            =   90
      TabIndex        =   40
      Top             =   536
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact Person"
      Height          =   240
      Index           =   10
      Left            =   90
      TabIndex        =   39
      Top             =   3390
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner"
      Height          =   240
      Index           =   6
      Left            =   90
      TabIndex        =   38
      Top             =   3028
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Tel. No."
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   37
      Top             =   1960
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Agreement"
      Height          =   240
      Index           =   3
      Left            =   90
      TabIndex        =   36
      Top             =   1248
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mode of Delivery"
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   35
      Top             =   892
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Company"
      Height          =   240
      Index           =   1
      Left            =   90
      TabIndex        =   34
      Top             =   180
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Area Address"
      Height          =   255
      Index           =   12
      Left            =   5130
      TabIndex        =   33
      Top             =   165
      Width           =   1185
   End
End
Attribute VB_Name = "frmSuppliersAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode
Public srcTextAdd           As TextBox 'Used in pop-up mode -> Display the customer address
Public srcTextCP            As TextBox 'Used in pop-up mode -> Display the customer contact person
Public srcTextDisc          As Object  'Used in pop-up mode -> Display the customer Discount (can be combo or textbox)

Dim cIRowCountBank          As Integer
Dim cIRowCountStock          As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset
Dim blnRemarks As Boolean

Private Sub DisplayForEditing()
    On Error GoTo errHandler
    
    With rs
        txtEntry(0).Text = .Fields("company")
        dcLocation.BoundText = .Fields![LocationID]
        cboDelivery.Text = .Fields("Delivery")
        cboFreightAgreement.Text = .Fields("FreightAgreement")
        cboFreightPeriod.Text = .Fields("FreightPeriod")
        txtEntry(1).Text = .Fields("Telephone")
        txtEntry(2).Text = .Fields("Owner")
        txtEntry(3).Text = .Fields("ContactPerson")
        txtEntry(4).Text = .Fields("Address1")
        txtEntry(5).Text = .Fields("Address2")
        txtEntry(6).Text = .Fields("Cellphone")
        txtEntry(11).Text = .Fields("Fax")
        txtEntry(12).Text = .Fields("CreditTerm")
        cboCreditOption.Text = IIf(IsNull(.Fields("CreditOption")), "", .Fields("CreditOption"))
        txtEntry(13).Text = .Fields("CreditLimit")
        txtEntry(14).Text = .Fields("Remarks")
        txtEntry(15).Text = toNumber(.Fields("GenDiscPercent"))
    End With
    
    'Display the details of Bank
    Dim rsVendorBank As New Recordset

    cIRowCountBank = 0
    
    rsVendorBank.CursorLocation = adUseClient
    rsVendorBank.Open "SELECT * FROM qry_Vendors_Bank WHERE VendorID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsVendorBank.RecordCount > 0 Then
        rsVendorBank.MoveFirst
        While Not rsVendorBank.EOF
          cIRowCountBank = cIRowCountBank + 1     'increment
            With GridBanks
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = rsVendorBank![Bank]
                    .TextMatrix(1, 2) = rsVendorBank![Branch]
                    .TextMatrix(1, 3) = rsVendorBank![AccountNo]
                    .TextMatrix(1, 4) = rsVendorBank![AccountName]
                    .TextMatrix(1, 5) = rsVendorBank![BankID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsVendorBank![Bank]
                    .TextMatrix(.Rows - 1, 2) = rsVendorBank![Branch]
                    .TextMatrix(.Rows - 1, 3) = rsVendorBank![AccountNo]
                    .TextMatrix(.Rows - 1, 4) = rsVendorBank![AccountName]
                    .TextMatrix(.Rows - 1, 5) = rsVendorBank![BankID]
                End If
            End With
            rsVendorBank.MoveNext
        Wend
        GridBanks.Row = 1
        GridBanks.ColSel = 5
        'Set fixed cols
        If State = adStateEditMode Then
            GridBanks.FixedRows = GridBanks.Row: 'GridBanks.SelectionMode = flexSelectionFree
            GridBanks.FixedCols = 1
        End If
    End If

    'Display the details of Stock
    Dim rsVendorStock As New Recordset

    cIRowCountStock = 0
    
    rsVendorStock.CursorLocation = adUseClient
    rsVendorStock.Open "SELECT * FROM qry_Vendors_Stocks WHERE VendorID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsVendorStock.RecordCount > 0 Then
        rsVendorStock.MoveFirst
        While Not rsVendorStock.EOF
          cIRowCountStock = cIRowCountStock + 1     'increment
            With GridStocks
                If .Rows = 2 And .TextMatrix(1, 6) = "" Then
                    .TextMatrix(1, 1) = rsVendorStock!Stock
                    .TextMatrix(1, 2) = rsVendorStock!DiscPercent
                    .TextMatrix(1, 3) = rsVendorStock!ExtDiscPercent
                    .TextMatrix(1, 4) = rsVendorStock!ExtDiscAmount
                    .TextMatrix(1, 5) = rsVendorStock!SupplierPrice
                    .TextMatrix(1, 6) = rsVendorStock!ExtPrice
                    .TextMatrix(1, 7) = rsVendorStock!StockID
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsVendorStock!Stock
                    .TextMatrix(.Rows - 1, 2) = rsVendorStock!DiscPercent
                    .TextMatrix(.Rows - 1, 3) = rsVendorStock!ExtDiscPercent
                    .TextMatrix(.Rows - 1, 4) = rsVendorStock!ExtDiscAmount
                    .TextMatrix(.Rows - 1, 5) = rsVendorStock!SupplierPrice
                    .TextMatrix(.Rows - 1, 6) = rsVendorStock!ExtPrice
                    .TextMatrix(.Rows - 1, 7) = rsVendorStock!StockID
                End If
            End With
            rsVendorStock.MoveNext
        Wend
        GridStocks.Row = 1
        GridStocks.ColSel = 7
        'Set fixed cols
        If State = adStateEditMode Then
            GridStocks.FixedRows = GridStocks.Row: 'GridStocks.SelectionMode = flexSelectionFree
            GridStocks.FixedCols = 1
        End If
    End If
    
    rsVendorBank.Close
    rsVendorStock.Close
    'Clear variables
    Set rsVendorBank = Nothing
    Set rsVendorStock = Nothing
    
    Exit Sub
errHandler:
  If err.Number = 94 Then
    Resume Next
  Else
    MsgBox "Error: " & err.Description
  End If
End Sub

Private Sub btnAddBank_Click()
    If nsdBank.Text = "" Or txtAcctNo.Text = "" Or txtAcctName.Text = "" Then nsdBank.SetFocus: Exit Sub

    Dim CurrRow As Integer
    Dim intBankID As Integer
    
    If nsdBank.BoundText = "" Then
        CurrRow = getFlexPos(GridBanks, 5, nsdBank.Tag)
        intBankID = nsdBank.Tag
    Else
        CurrRow = getFlexPos(GridBanks, 5, nsdBank.BoundText)
        intBankID = nsdBank.BoundText
    End If

    'Add to GridBanks
    With GridBanks
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                .TextMatrix(1, 1) = nsdBank.Text
                .TextMatrix(1, 2) = txtBranch.Text
                .TextMatrix(1, 3) = txtAcctNo.Text
                .TextMatrix(1, 4) = txtAcctName.Text
                .TextMatrix(1, 5) = intBankID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdBank.Text
                .TextMatrix(.Rows - 1, 2) = txtBranch.Text
                .TextMatrix(.Rows - 1, 3) = txtAcctNo.Text
                .TextMatrix(.Rows - 1, 4) = txtAcctName.Text
                .TextMatrix(.Rows - 1, 5) = intBankID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCountBank = cIRowCountBank + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = nsdBank.Text
                .TextMatrix(CurrRow, 2) = txtBranch.Text
                .TextMatrix(CurrRow, 3) = txtAcctNo.Text
                .TextMatrix(CurrRow, 4) = txtAcctName.Text
                .TextMatrix(CurrRow, 5) = intBankID
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .ColSel = 5
        'Display a remove button
        GridBanks_Click
    End With
End Sub

Private Sub btnRemoveBank_Click()
    'Remove selected load product
    With GridBanks
        'Update the record count
        cIRowCountBank = cIRowCountBank - 1
        
        If .Rows = 2 Then .Rows = .Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemoveBank.Visible = False
    GridBanks_Click
End Sub

Private Sub btnRemoveStock_Click()
    'Remove selected load product
    With GridStocks
        'Update the record count
        cIRowCountStock = cIRowCountStock - 1
        
        If .Rows = 2 Then .Rows = .Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemoveStock.Visible = False
    GridStocks_Click
End Sub

Private Sub cmdAddStock_Click()
    If nsdStock.Text = "" Or txtExtDiscPercent.Text = "" Or txtExtDiscAmount.Text = "" Or txtUnitPrice.Text = "" Then nsdBank.SetFocus: Exit Sub

    Dim CurrRow As Integer
    Dim intStockID As Integer
    
    If nsdStock.BoundText = "" Then
        CurrRow = getFlexPos(GridStocks, 7, nsdStock.Tag)
        intStockID = nsdStock.Tag
    Else
        CurrRow = getFlexPos(GridStocks, 7, nsdStock.BoundText)
        intStockID = nsdStock.BoundText
    End If

    'Add to GridStocks
    With GridStocks
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 7) = "" Then
                .TextMatrix(1, 1) = nsdStock.Text
                .TextMatrix(1, 2) = txtDiscPercent.Text
                .TextMatrix(1, 3) = txtExtDiscPercent.Text
                .TextMatrix(1, 4) = txtExtDiscAmount.Text
                .TextMatrix(1, 5) = txtUnitPrice.Text
                .TextMatrix(1, 6) = txtExtPrice.Text
                .TextMatrix(1, 7) = intStockID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.Text
                .TextMatrix(.Rows - 1, 2) = txtDiscPercent.Text
                .TextMatrix(.Rows - 1, 3) = txtExtDiscPercent.Text
                .TextMatrix(.Rows - 1, 4) = txtExtDiscAmount.Text
                .TextMatrix(.Rows - 1, 5) = txtUnitPrice.Text
                .TextMatrix(.Rows - 1, 6) = txtExtPrice.Text
                .TextMatrix(.Rows - 1, 7) = intStockID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCountStock = cIRowCountStock + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = nsdStock.Text
                .TextMatrix(CurrRow, 2) = txtDiscPercent.Text
                .TextMatrix(CurrRow, 3) = txtExtDiscPercent.Text
                .TextMatrix(CurrRow, 4) = txtExtDiscAmount.Text
                .TextMatrix(CurrRow, 5) = txtUnitPrice.Text
                .TextMatrix(CurrRow, 6) = txtExtPrice.Text
                .TextMatrix(CurrRow, 7) = intStockID
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .ColSel = 5
        'Display a remove button
        GridStocks_Click
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    If is_empty(txtEntry(0), True) = True Then
      MsgBox "The field 'Company' is required. Please check it!", vbExclamation
      Exit Sub
    End If
    
    If Trim(dcLocation.Text) = "" Then
      MsgBox "The field 'Category' is required. Please check it!", vbExclamation
      Exit Sub
    End If
    
    If Trim(cboDelivery) = "" Then
      MsgBox "The field 'Mode of Delivery' is required. Please check it!", vbExclamation
      Exit Sub
    End If
    
    If Trim(cboFreightAgreement) = "" Then
      MsgBox "The field 'Frieght Payment Agreement' is required.Please check it!", vbExclamation
      Exit Sub
    End If
    
    If Trim(cboFreightPeriod) = "" Then
      MsgBox "The field 'Frieght Payment Period' is required.Please check it!", vbExclamation
      Exit Sub
    End If
    
    If Trim(cboCreditOption) = "" Then
      MsgBox "The field 'Credit Option' is required. Please check it!", vbExclamation
      Exit Sub
    End If
    
    CN.BeginTrans
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        
        rs.Fields("VendorID") = PK
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("company") = txtEntry(0).Text
        .Fields("LocationID") = dcLocation.BoundText
        .Fields("delivery") = cboDelivery.Text
        .Fields("FreightAgreement") = cboFreightAgreement.Text
        .Fields("FreightPeriod") = cboFreightPeriod.Text
        .Fields("Telephone") = txtEntry(1).Text
        .Fields("Owner") = txtEntry(2).Text
        .Fields("ContactPerson") = txtEntry(3).Text
        .Fields("Address1") = txtEntry(4).Text
        .Fields("Address2") = txtEntry(5).Text
        .Fields("Cellphone") = txtEntry(6).Text
        .Fields("Fax") = txtEntry(11).Text
        .Fields("CreditTerm") = txtEntry(12).Text
        .Fields("CreditOption") = cboCreditOption.Text
        .Fields("CreditLimit") = toNumber(txtEntry(13).Text)
        .Fields("Remarks") = txtEntry(14).Text
        .Fields("GenDiscPercent") = toNumber(txtEntry(15).Text)
        .Update
    End With
    
    Dim rsVendorBank As New Recordset

    rsVendorBank.CursorLocation = adUseClient
    rsVendorBank.Open "SELECT * FROM Vendors_Bank WHERE VendorID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItemsBank
    
    Dim cBank As Integer
    
    With GridBanks
        'Save the details of the records
        For cBank = 1 To cIRowCountBank
            .Row = cBank
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNewBank:
                rsVendorBank.AddNew

                rsVendorBank![VendorID] = PK
                rsVendorBank![BankID] = toNumber(.TextMatrix(cBank, 5))
                rsVendorBank![AccountNo] = .TextMatrix(cBank, 3)
                rsVendorBank![AccountName] = .TextMatrix(cBank, 4)

                rsVendorBank.Update
            ElseIf State = adStateEditMode Then
                rsVendorBank.Filter = "BankID = " & toNumber(.TextMatrix(cBank, 5))
            
                If rsVendorBank.RecordCount = 0 Then GoTo AddNewBank

                rsVendorBank![VendorID] = PK
                rsVendorBank![BankID] = toNumber(.TextMatrix(cBank, 5))
                rsVendorBank![AccountNo] = .TextMatrix(cBank, 3)
                rsVendorBank![AccountName] = .TextMatrix(cBank, 4)

                rsVendorBank.Update
            End If

        Next cBank
    End With
   
    '-------------Stocks--------------------
    
    Set rsVendorBank = Nothing
    
    Dim rsVendorStock As New Recordset
    
    'save vendorstock
    rsVendorStock.CursorLocation = adUseClient
    rsVendorStock.Open "SELECT * FROM Vendors_Stocks WHERE VendorId = " & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItemsStock
    
    Dim cStock As Integer
    
    With GridStocks
        'Save the details of the records
        For cStock = 1 To cIRowCountStock
            .Row = cStock
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNewStock:
                rsVendorStock.AddNew

                rsVendorStock![VendorID] = PK
                rsVendorStock![StockID] = toNumber(.TextMatrix(cStock, 7))
                rsVendorStock![DiscPercent] = .TextMatrix(cStock, 2)
                rsVendorStock![ExtDiscPercent] = .TextMatrix(cStock, 3)
                rsVendorStock![ExtDiscAmount] = .TextMatrix(cStock, 4)
                rsVendorStock![Price] = .TextMatrix(cStock, 5)
                rsVendorStock![ExtPrice] = .TextMatrix(cStock, 6)

                rsVendorStock.Update
            ElseIf State = adStateEditMode Then
                rsVendorStock.Filter = "StockID = " & toNumber(.TextMatrix(cStock, 7))
            
                If rsVendorStock.RecordCount = 0 Then GoTo AddNewStock

                rsVendorStock![VendorID] = PK
                rsVendorStock![StockID] = toNumber(.TextMatrix(cStock, 7))
                rsVendorStock![DiscPercent] = .TextMatrix(cStock, 2)
                rsVendorStock![ExtDiscPercent] = .TextMatrix(cStock, 3)
                rsVendorStock![ExtDiscAmount] = .TextMatrix(cStock, 4)
                rsVendorStock![Price] = .TextMatrix(cStock, 5)
                rsVendorStock![ExtPrice] = .TextMatrix(cStock, 6)

                rsVendorStock.Update
            End If

        Next cStock
    End With
    
    'Clear variables
    cBank = 0
    cStock = 0
    
    HaveAction = True
    
    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
        Else
          Unload Me
        End If
    ElseIf State = adStatePopupMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        Unload Me
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
    
    CN.CommitTrans
    
    Exit Sub
err:
  
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmSupplier" & vbCr _
  & "Sub: cmdSave_Click", vbExclamation, "Error"
  If err.Number = -2147217887 Then Resume Next
        
  CN.RollbackTrans
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(rs.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("LastUserFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: " & tDate2 & vbCrLf & _
           "Modified By: " & tUser2, vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tDate2 = vbNullString
    tUser1 = vbNullString
    tUser2 = vbNullString
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And blnRemarks = False Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    InitGrid

    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Vendors WHERE VendorID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Vendors_Location", "Location", dcLocation, "LocationID", True
    
    'Check the form state
    InitNSD
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
      
        GeneratePK
    Else
      Caption = "Edit Entry"
      DisplayForEditing
    End If
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Vendors")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmSuppliers.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmSuppliersAE = Nothing
End Sub

Private Sub GridBanks_Click()
    With GridBanks
        nsdBank.Text = .TextMatrix(.RowSel, 1)
        nsdBank.Tag = .TextMatrix(.RowSel, 5) 'Add tag coz boundtext is empty
        txtBranch.Text = .TextMatrix(.RowSel, 2)
        txtAcctNo.Text = .TextMatrix(.RowSel, 3)
        txtAcctName.Text = .TextMatrix(.RowSel, 4)
    
        If GridBanks.Rows = 2 And GridBanks.TextMatrix(1, 5) = "" Then '10 = StockID
            btnRemoveBank.Visible = False
        Else
            btnRemoveBank.Visible = True
            btnRemoveBank.Top = (GridBanks.CellTop + GridBanks.Top) - 20
            btnRemoveBank.Left = GridBanks.Left + 50
        End If
    End With
End Sub

Private Sub GridStocks_Click()
    With GridStocks
        nsdStock.Text = .TextMatrix(.RowSel, 1)
        nsdStock.Tag = .TextMatrix(.RowSel, 7) 'Add tag coz boundtext is empty
        txtDiscPercent.Text = .TextMatrix(.RowSel, 2)
        txtExtDiscPercent.Text = .TextMatrix(.RowSel, 3)
        txtExtDiscAmount.Text = .TextMatrix(.RowSel, 4)
        txtUnitPrice.Text = .TextMatrix(.RowSel, 5)
        txtExtPrice.Text = .TextMatrix(.RowSel, 6)
    
        If GridStocks.Rows = 2 And GridStocks.TextMatrix(1, 5) = "" Then '10 = StockID
            btnRemoveStock.Visible = False
        Else
            btnRemoveStock.Visible = True
            btnRemoveStock.Top = (GridStocks.CellTop + GridStocks.Top) - 20
            btnRemoveStock.Left = GridStocks.Left + 50
        End If
    End With
End Sub

Private Sub GridStocks_DblClick()
    With frmLandedCost
        .intProductID = GridStocks.TextMatrix(GridStocks.RowSel, 7) 'StockID
        .intSupplierID = PK
        
        .show 1
    End With
End Sub

Private Sub nsdBank_Change()
    txtBranch.Text = nsdBank.getSelValueAt(3)
End Sub

Private Sub nsdStock_Change()
  txtUnitPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Selling Price
End Sub

Private Sub txtExtDiscAmount_GotFocus()
    HLText txtExtDiscAmount
End Sub

Private Sub txtExtDiscAmount_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtextDiscPercent_GotFocus()
    HLText txtExtDiscPercent
End Sub

Private Sub txtextDiscPercent_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 14 Then
        blnRemarks = True
        Exit Sub
    Else
        blnRemarks = False
    End If
    
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 12 Or Index = 13 Or Index = 15 Then KeyAscii = isNumber(KeyAscii)
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCountBank = 0
    With GridBanks
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 6
        .ColSel = 5
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1400
        .ColWidth(2) = 1500
        .ColWidth(3) = 1400
        .ColWidth(4) = 1500
        .ColWidth(5) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Bank"
        .TextMatrix(0, 2) = "Branch"
        .TextMatrix(0, 3) = "Acct. No."
        .TextMatrix(0, 4) = "Acct. Name"
        .TextMatrix(0, 5) = "Bank ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
        .ColAlignment(3) = vbLeftJustify
        .ColAlignment(4) = vbRightJustify
    End With
    
    cIRowCountStock = 0
    With GridStocks
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 8
        .ColSel = 7
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 3400
        .ColWidth(2) = 1000
        .ColWidth(3) = 1000
        .ColWidth(4) = 1400
        .ColWidth(5) = 1400
        .ColWidth(6) = 1000
        .ColWidth(7) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product"
        .TextMatrix(0, 2) = "Disc (%)"
        .TextMatrix(0, 3) = "Ext Disc (%)"
        .TextMatrix(0, 4) = "Ext Disc (Amt)"
        .TextMatrix(0, 5) = "Price"
        .TextMatrix(0, 6) = "Ext Price"
        .TextMatrix(0, 7) = "StockID"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbRightJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbRightJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbRightJustify
'        .ColAlignment(6) = vbRightJustify
    End With
End Sub

Private Sub InitNSD()
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        .AddColumn "Supplier Price", 1500
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,SupplierPrice,StockID"
        .sqlTables = "qry_Stock_Unit"
        .sqlwCondition = "qry_Stock_Unit.Order=1"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Stocks"
    End With

    'For Bank
    With nsdBank
        .ClearColumn
        .AddColumn "Bank ID", 1794.89
        .AddColumn "Bank", 2264.88
        .AddColumn "Branch", 2670.23
        .Connection = CN.ConnectionString
        
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

Private Sub DeleteItemsBank()
    Dim CurrRow As Integer
    Dim RSBank As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSBank.CursorLocation = adUseClient
    RSBank.Open "SELECT * FROM Vendors_Bank WHERE VendorID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSBank.RecordCount > 0 Then
        RSBank.MoveFirst
        While Not RSBank.EOF
            CurrRow = getFlexPos(GridBanks, 5, RSBank!BankID)
        
            'Add to GridBanks
            With GridBanks
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Vendors_Bank", "PK", "", True, RSBank!PK
                End If
            End With
            RSBank.MoveNext
        Wend
    End If
End Sub

Private Sub DeleteItemsStock()
    Dim CurrRow As Integer
    Dim RSStock As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSStock.CursorLocation = adUseClient
    RSStock.Open "SELECT * FROM Vendors_Stocks WHERE VendorID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSStock.RecordCount > 0 Then
        RSStock.MoveFirst
        While Not RSStock.EOF
            CurrRow = getFlexPos(GridStocks, 7, RSStock!StockID)
        
            'Add to GridBanks
            With GridStocks
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Vendors_Stocks", "VendorStockID", "", True, RSStock!VendorStockID
                End If
            End With
            RSStock.MoveNext
        Wend
    End If
End Sub

