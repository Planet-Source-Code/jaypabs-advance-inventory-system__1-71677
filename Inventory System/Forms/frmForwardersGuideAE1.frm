VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmForwardersGuideAE1 
   BorderStyle     =   0  'None
   ClientHeight    =   9930
   ClientLeft      =   1050
   ClientTop       =   570
   ClientWidth     =   12510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9930
   ScaleWidth      =   12510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10830
      Locked          =   -1  'True
      TabIndex        =   118
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6480
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
      Index           =   1
      Left            =   10830
      Locked          =   -1  'True
      TabIndex        =   117
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   6180
      Width           =   1425
   End
   Begin VB.TextBox txtNet 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   10830
      Locked          =   -1  'True
      TabIndex        =   116
      TabStop         =   0   'False
      Text            =   "0.00"
      Top             =   7530
      Width           =   1425
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2985
      Left            =   210
      TabIndex        =   62
      Top             =   6150
      Width           =   7605
      _ExtentX        =   13414
      _ExtentY        =   5265
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "Transportation Cost"
      TabPicture(0)   =   "frmForwardersGuideAE1.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label33"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label32"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label31"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label30"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label29"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label28"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label27"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label26"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label25"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label24"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label23"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label22"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label21"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label20"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label19"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label17"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label16"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label9"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label8"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label6"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "dtp7"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "dtp6"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "dtp5"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "dtp4"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "dtp3"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "dtp2"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "dtp1"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtAmount7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtOR7"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "txtAmount6"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "txtOR6"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "txtAmount5"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "txtOR5"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtAmount4"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtOR4"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtAmount3"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txtOR3"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "txtAmount2"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "txtOR2"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtAmount1"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtOR1"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).ControlCount=   42
      TabCaption(1)   =   "Payment"
      TabPicture(1)   =   "frmForwardersGuideAE1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "txtCheckNo"
      Tab(1).Control(2)=   "txtReceipt"
      Tab(1).Control(3)=   "cboAmbot"
      Tab(1).Control(4)=   "dtpReceiptDate"
      Tab(1).Control(5)=   "Label35"
      Tab(1).Control(6)=   "Label36"
      Tab(1).Control(7)=   "Label37"
      Tab(1).Control(8)=   "Label38"
      Tab(1).ControlCount=   9
      Begin VB.Frame Frame1 
         Caption         =   "Account Type"
         Height          =   855
         Left            =   -74790
         TabIndex        =   113
         Top             =   420
         Width           =   3285
         Begin VB.OptionButton optSeperate 
            Caption         =   "Seperate"
            Height          =   285
            Left            =   1470
            TabIndex        =   115
            Top             =   360
            Width           =   1125
         End
         Begin VB.OptionButton optCombined 
            Caption         =   "Combined"
            Height          =   285
            Left            =   150
            TabIndex        =   114
            Top             =   360
            Value           =   -1  'True
            Width           =   1035
         End
      End
      Begin VB.TextBox txtCheckNo 
         Height          =   315
         Left            =   -73320
         TabIndex        =   107
         Top             =   2490
         Width           =   1815
      End
      Begin VB.TextBox txtReceipt 
         Height          =   315
         Left            =   -73320
         TabIndex        =   106
         Top             =   2130
         Width           =   1815
      End
      Begin VB.ComboBox cboAmbot 
         Height          =   315
         ItemData        =   "frmForwardersGuideAE1.frx":0038
         Left            =   -73320
         List            =   "frmForwardersGuideAE1.frx":0042
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   1770
         Width           =   1815
      End
      Begin VB.TextBox txtOR1 
         Height          =   315
         Left            =   4590
         TabIndex        =   76
         Top             =   390
         Width           =   795
      End
      Begin VB.TextBox txtAmount1 
         Height          =   315
         Left            =   6240
         TabIndex        =   75
         Text            =   "0.00"
         Top             =   390
         Width           =   1185
      End
      Begin VB.TextBox txtOR2 
         Height          =   315
         Left            =   4590
         TabIndex        =   74
         Top             =   750
         Width           =   795
      End
      Begin VB.TextBox txtAmount2 
         Height          =   315
         Left            =   6240
         TabIndex        =   73
         Text            =   "0.00"
         Top             =   750
         Width           =   1185
      End
      Begin VB.TextBox txtOR3 
         Height          =   315
         Left            =   4590
         TabIndex        =   72
         Top             =   1110
         Width           =   795
      End
      Begin VB.TextBox txtAmount3 
         Height          =   315
         Left            =   6240
         TabIndex        =   71
         Text            =   "0.00"
         Top             =   1110
         Width           =   1185
      End
      Begin VB.TextBox txtOR4 
         Height          =   315
         Left            =   4590
         TabIndex        =   70
         Top             =   1470
         Width           =   795
      End
      Begin VB.TextBox txtAmount4 
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   69
         Text            =   "0.00"
         Top             =   1470
         Width           =   1185
      End
      Begin VB.TextBox txtOR5 
         Height          =   315
         Left            =   4590
         TabIndex        =   68
         Top             =   1830
         Width           =   795
      End
      Begin VB.TextBox txtAmount5 
         Height          =   315
         Left            =   6240
         TabIndex        =   67
         Text            =   "0.00"
         Top             =   1830
         Width           =   1185
      End
      Begin VB.TextBox txtOR6 
         Height          =   315
         Left            =   4590
         TabIndex        =   66
         Top             =   2190
         Width           =   795
      End
      Begin VB.TextBox txtAmount6 
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   65
         Text            =   "0.00"
         Top             =   2190
         Width           =   1185
      End
      Begin VB.TextBox txtOR7 
         Height          =   315
         Left            =   4590
         TabIndex        =   64
         Top             =   2550
         Width           =   795
      End
      Begin VB.TextBox txtAmount7 
         Height          =   315
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   63
         Text            =   "0.00"
         Top             =   2550
         Width           =   1185
      End
      Begin MSComCtl2.DTPicker dtp1 
         Height          =   315
         Left            =   2160
         TabIndex        =   77
         Top             =   390
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp2 
         Height          =   315
         Left            =   2160
         TabIndex        =   78
         Top             =   750
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp3 
         Height          =   315
         Left            =   2160
         TabIndex        =   79
         Top             =   1110
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp4 
         Height          =   315
         Left            =   2160
         TabIndex        =   80
         Top             =   1470
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp5 
         Height          =   315
         Left            =   2160
         TabIndex        =   81
         Top             =   1830
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp6 
         Height          =   315
         Left            =   2160
         TabIndex        =   82
         Top             =   2190
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtp7 
         Height          =   315
         Left            =   2160
         TabIndex        =   83
         Top             =   2550
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CheckBox        =   -1  'True
         CustomFormat    =   "MMM-dd- yyyy"
         DateIsNull      =   -1  'True
         Format          =   44761091
         CurrentDate     =   39039
      End
      Begin MSComCtl2.DTPicker dtpReceiptDate 
         Height          =   315
         Left            =   -73320
         TabIndex        =   108
         Top             =   1410
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "MMM-dd-yyyy"
         Format          =   44761091
         CurrentDate     =   38989
      End
      Begin VB.Label Label35 
         Caption         =   "Check #"
         Height          =   225
         Left            =   -74760
         TabIndex        =   112
         Top             =   2490
         Width           =   735
      End
      Begin VB.Label Label36 
         Caption         =   "Date"
         Height          =   225
         Left            =   -74760
         TabIndex        =   111
         Top             =   1410
         Width           =   735
      End
      Begin VB.Label Label37 
         Caption         =   "Receipt #"
         Height          =   225
         Left            =   -74760
         TabIndex        =   110
         Top             =   2130
         Width           =   735
      End
      Begin VB.Label Label38 
         Caption         =   "Payment Type"
         Height          =   225
         Left            =   -74760
         TabIndex        =   109
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Mla. Trucking Date"
         Height          =   225
         Left            =   90
         TabIndex        =   104
         Top             =   390
         Width           =   1395
      End
      Begin VB.Label Label6 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   103
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   102
         Top             =   390
         Width           =   675
      End
      Begin VB.Label Label9 
         Caption         =   "Mla. Arrastre Date"
         Height          =   225
         Left            =   90
         TabIndex        =   101
         Top             =   750
         Width           =   1395
      End
      Begin VB.Label Label16 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   100
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   99
         Top             =   750
         Width           =   675
      End
      Begin VB.Label Label19 
         Caption         =   "Mla. Wfg. Fee Date"
         Height          =   225
         Left            =   90
         TabIndex        =   98
         Top             =   1110
         Width           =   1395
      End
      Begin VB.Label Label20 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   97
         Top             =   1110
         Width           =   735
      End
      Begin VB.Label Label21 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   96
         Top             =   1110
         Width           =   675
      End
      Begin VB.Label Label22 
         Caption         =   "Freight Date"
         Height          =   225
         Left            =   90
         TabIndex        =   95
         Top             =   1470
         Width           =   1395
      End
      Begin VB.Label Label23 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   94
         Top             =   1470
         Width           =   735
      End
      Begin VB.Label Label24 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   93
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label Label25 
         Caption         =   "Local Arrastre Date"
         Height          =   225
         Left            =   90
         TabIndex        =   92
         Top             =   1830
         Width           =   1395
      End
      Begin VB.Label Label26 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   91
         Top             =   1830
         Width           =   735
      End
      Begin VB.Label Label27 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   90
         Top             =   1830
         Width           =   675
      End
      Begin VB.Label Label28 
         Caption         =   "Local Trucking Date"
         Height          =   225
         Left            =   90
         TabIndex        =   89
         Top             =   2190
         Width           =   1845
      End
      Begin VB.Label Label29 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   88
         Top             =   2190
         Width           =   735
      End
      Begin VB.Label Label30 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   87
         Top             =   2190
         Width           =   675
      End
      Begin VB.Label Label31 
         Caption         =   "Sidewalk Handling Date"
         Height          =   225
         Left            =   90
         TabIndex        =   86
         Top             =   2550
         Width           =   1845
      End
      Begin VB.Label Label32 
         Caption         =   "O.R.#"
         Height          =   225
         Left            =   4080
         TabIndex        =   85
         Top             =   2550
         Width           =   735
      End
      Begin VB.Label Label33 
         Caption         =   "Amount"
         Height          =   225
         Left            =   5520
         TabIndex        =   84
         Top             =   2550
         Width           =   675
      End
   End
   Begin VB.TextBox txtTruckNo 
      Height          =   314
      Left            =   11010
      TabIndex        =   58
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtShippingGuideNo 
      Height          =   314
      Left            =   1800
      TabIndex        =   56
      Top             =   1560
      Width           =   1545
   End
   Begin MSDataListLib.DataCombo dcClass 
      Height          =   315
      Left            =   1800
      TabIndex        =   55
      Top             =   2280
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtShip 
      Height          =   314
      Left            =   8910
      TabIndex        =   51
      Top             =   1560
      Width           =   3315
   End
   Begin VB.TextBox txtDate 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   44
      Top             =   990
      Width           =   1905
   End
   Begin VB.TextBox txtAddress 
      Height          =   285
      Left            =   6360
      Locked          =   -1  'True
      TabIndex        =   43
      Top             =   660
      Width           =   4005
   End
   Begin VB.TextBox txtVanNo 
      Height          =   314
      Left            =   11010
      TabIndex        =   41
      Top             =   1920
      Width           =   1185
   End
   Begin VB.TextBox txtVoyageNo 
      Height          =   314
      Left            =   8910
      TabIndex        =   36
      Top             =   1920
      Width           =   1095
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   180
      ScaleHeight     =   720
      ScaleWidth      =   12120
      TabIndex        =   9
      Top             =   3060
      Width           =   12120
      Begin VB.TextBox txtFreight 
         Height          =   315
         Left            =   10530
         TabIndex        =   52
         Top             =   390
         Width           =   735
      End
      Begin VB.TextBox txtNetAmount 
         Height          =   315
         Left            =   9420
         Locked          =   -1  'True
         TabIndex        =   49
         Text            =   "0.00"
         Top             =   390
         Width           =   1065
      End
      Begin VB.TextBox txtDisc 
         Height          =   315
         Left            =   8820
         TabIndex        =   47
         Text            =   "0"
         Top             =   390
         Width           =   555
      End
      Begin VB.TextBox txtGross 
         Height          =   315
         Index           =   0
         Left            =   7710
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "0.00"
         Top             =   360
         Width           =   1035
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   4170
         TabIndex        =   40
         Top             =   360
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.TextBox txtOQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3090
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "0"
         Top             =   360
         Width           =   480
      End
      Begin VB.ComboBox cboClass 
         Height          =   315
         ItemData        =   "frmForwardersGuideAE1.frx":0059
         Left            =   6270
         List            =   "frmForwardersGuideAE1.frx":005B
         Style           =   2  'Dropdown List
         TabIndex        =   35
         Top             =   360
         Width           =   1395
      End
      Begin VB.TextBox txtStock 
         Height          =   315
         Left            =   0
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox txtValue 
         Height          =   315
         Left            =   5160
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   360
         Width           =   1035
      End
      Begin VB.TextBox txtRQty 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3540
         TabIndex        =   11
         Text            =   "0"
         Top             =   360
         Width           =   510
      End
      Begin VB.CommandButton btnUpdate 
         Caption         =   "Update"
         Enabled         =   0   'False
         Height          =   315
         Left            =   11340
         TabIndex        =   10
         Top             =   375
         Width           =   720
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Freight %"
         Height          =   240
         Index           =   7
         Left            =   10560
         TabIndex        =   53
         Top             =   60
         Width           =   720
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Net Amount"
         Height          =   240
         Index           =   6
         Left            =   9450
         TabIndex        =   50
         Top             =   60
         Width           =   1050
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Disc %"
         Height          =   240
         Index           =   5
         Left            =   8850
         TabIndex        =   48
         Top             =   60
         Width           =   510
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Gross"
         Height          =   240
         Index           =   4
         Left            =   7710
         TabIndex        =   46
         Top             =   30
         Width           =   1050
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "O-Qty"
         Height          =   240
         Index           =   3
         Left            =   3090
         TabIndex        =   39
         Top             =   90
         Width           =   480
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Loose Cargo"
         Height          =   240
         Index           =   17
         Left            =   6270
         TabIndex        =   18
         Top             =   60
         Width           =   1050
      End
      Begin VB.Label Labels 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "R-Qty"
         Height          =   240
         Index           =   10
         Left            =   3600
         TabIndex        =   17
         Top             =   60
         Width           =   510
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
         Height          =   240
         Index           =   9
         Left            =   5190
         TabIndex        =   16
         Top             =   90
         Width           =   990
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Product/ Item"
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
         TabIndex        =   15
         Top             =   90
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   4170
         TabIndex        =   14
         Top             =   90
         Width           =   900
      End
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   240
      Picture         =   "frmForwardersGuideAE1.frx":005D
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Remove"
      Top             =   4170
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   9300
      TabIndex        =   7
      Top             =   9210
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   10845
      TabIndex        =   6
      Top             =   9210
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   210
      TabIndex        =   5
      Top             =   9210
      Width           =   1755
   End
   Begin VB.TextBox txtPONo 
      Height          =   285
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   660
      Width           =   1905
   End
   Begin VB.TextBox txtBLNo 
      Height          =   314
      Left            =   8910
      TabIndex        =   3
      Top             =   2280
      Width           =   1095
   End
   Begin VB.TextBox txtDRNo 
      Height          =   314
      Left            =   6000
      TabIndex        =   2
      Top             =   1920
      Width           =   1785
   End
   Begin VB.TextBox txtSupplier 
      Height          =   314
      Left            =   1800
      TabIndex        =   1
      Top             =   990
      Width           =   3075
   End
   Begin VB.CommandButton CmdReturn 
      Caption         =   "Receive Items"
      Height          =   315
      Left            =   7890
      TabIndex        =   0
      Top             =   9210
      Visible         =   0   'False
      Width           =   1335
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   7920
      TabIndex        =   19
      Top             =   9090
      Width           =   4260
      _ExtentX        =   21114
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2280
      Left            =   180
      TabIndex        =   20
      Top             =   3780
      Width           =   12075
      _ExtentX        =   21299
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
   Begin ctrlNSDataCombo.NSDataCombo nsdShippingCo 
      Height          =   315
      Left            =   1800
      TabIndex        =   21
      Top             =   1920
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
   Begin MSComCtl2.DTPicker dtpShippingDate 
      Height          =   315
      Left            =   6000
      TabIndex        =   22
      Top             =   2280
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44761091
      CurrentDate     =   38989
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdLocal 
      Height          =   315
      Left            =   6000
      TabIndex        =   60
      Top             =   1560
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
   Begin VB.Label Label48 
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
      Left            =   8730
      TabIndex        =   121
      Top             =   6510
      Width           =   2040
   End
   Begin VB.Label Label4 
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
      Left            =   8730
      TabIndex        =   120
      Top             =   6210
      Width           =   2040
   End
   Begin VB.Label Label3 
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
      Left            =   8730
      TabIndex        =   119
      Top             =   7560
      Width           =   2040
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000000&
      BorderWidth     =   2
      X1              =   10500
      X2              =   12210
      Y1              =   7470
      Y2              =   7470
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
      TabIndex        =   61
      Top             =   1560
      Width           =   705
   End
   Begin VB.Label Label47 
      Alignment       =   1  'Right Justify
      Caption         =   "Truck No."
      Height          =   255
      Left            =   9540
      TabIndex        =   59
      Top             =   2340
      Width           =   1365
   End
   Begin VB.Label Label46 
      Caption         =   "Shipping Guide No."
      Height          =   255
      Left            =   210
      TabIndex        =   57
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label45 
      Caption         =   "Class"
      Height          =   255
      Left            =   180
      TabIndex        =   54
      Top             =   2280
      Width           =   1485
   End
   Begin VB.Label Label39 
      Alignment       =   1  'Right Justify
      Caption         =   "Van No."
      Height          =   255
      Left            =   9510
      TabIndex        =   42
      Top             =   1950
      Width           =   1365
   End
   Begin VB.Label Label1 
      Caption         =   "Voyage No."
      Height          =   225
      Left            =   8010
      TabIndex        =   37
      Top             =   1950
      Width           =   885
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   180
      X2              =   12240
      Y1              =   2700
      Y2              =   2700
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   12240
      Y1              =   2730
      Y2              =   2730
   End
   Begin VB.Label Label10 
      Caption         =   "Label10"
      Height          =   525
      Left            =   5070
      TabIndex        =   34
      Top             =   4260
      Width           =   1245
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
      Left            =   210
      TabIndex        =   33
      Top             =   120
      Width           =   4905
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   " Date"
      Height          =   225
      Index           =   1
      Left            =   4920
      TabIndex        =   32
      Top             =   990
      Width           =   1275
   End
   Begin VB.Label Labels 
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
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   31
      Top             =   1890
      Width           =   1515
   End
   Begin VB.Label Label15 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   225
      Left            =   4920
      TabIndex        =   30
      Top             =   645
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "PO No."
      Height          =   225
      Left            =   210
      TabIndex        =   29
      Top             =   660
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   0
      X1              =   150
      X2              =   12270
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   0
      X1              =   150
      X2              =   12270
      Y1              =   1410
      Y2              =   1410
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      Caption         =   "Ship"
      Height          =   255
      Left            =   7410
      TabIndex        =   28
      Top             =   1560
      Width           =   1365
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "B.L. No."
      Height          =   255
      Left            =   7470
      TabIndex        =   27
      Top             =   2310
      Width           =   1365
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   225
      Left            =   5040
      TabIndex        =   26
      Top             =   2310
      Width           =   795
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "D.R. No."
      Height          =   225
      Left            =   4470
      TabIndex        =   25
      Top             =   1950
      Width           =   1425
   End
   Begin VB.Label Label18 
      Caption         =   "Supplier"
      Height          =   225
      Left            =   210
      TabIndex        =   24
      Top             =   1020
      Width           =   1245
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Items"
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
      Left            =   270
      TabIndex        =   23
      Top             =   2820
      Width           =   915
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   90
      Top             =   90
      Width           =   12285
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2820
      Width           =   12030
   End
End
Attribute VB_Name = "frmForwardersGuideAE1"
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

Dim cCostPerPackage         As Double
Dim cTotalAmount            As Double
Dim cTotalTranspoCost       As Double

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim intQtyOld               As Integer 'Allowed value for receive qty
Dim dblLoose                As Double   'sum of all loose cargos




Private Sub btnUpdate_Click()
    Dim CurrRow As Integer

    CurrRow = getFlexPos(Grid, 10, Grid.TextMatrix(Grid.RowSel, 10))

    'validate the entry
    If txtRQty.Text = "0" Or txtValue.Text = "0.00" Or (cboClass.Text = "" And dcClass.Text = "Loose Cargo") Then Exit Sub
    If toNumber(txtOQty.Text) < toNumber(txtRQty.Text) Then
      MsgBox "Shipped Qty is greater than Ordered Qty.", vbExclamation
      Exit Sub
    End If
    
    'Add to grid
    With Grid
        .Row = CurrRow
        
        'If dcClass.Text = "Loose Cargo" Then
        '  dblLoose = dblLoose + GetFreight(nsdShippingCo.Text, cboClass.Text)
        '  txtAmount4.Text = toMoney(dblLoose)
        'Else
        '
        'End If
        
        
        .TextMatrix(CurrRow, 3) = txtRQty.Text
        .TextMatrix(CurrRow, 4) = dcUnit.Text
        .TextMatrix(CurrRow, 5) = toMoney(txtValue.Text)
        .TextMatrix(CurrRow, 6) = cboClass.Text
        .TextMatrix(CurrRow, 7) = toMoney(txtGross(0).Text)
        .TextMatrix(CurrRow, 8) = txtDisc.Text
        .TextMatrix(CurrRow, 9) = toMoney(txtNetAmount.Text)
        .TextMatrix(CurrRow, 10) = toNumber(txtFreight.Text)
        
        'compute total amount
        Dim i As Integer
        txtTotal.Text = 0
        For i = 1 To .Rows - 1
          txtTotal.Text = toMoney(txtTotal.Text) + toNumber(.TextMatrix(1, 9))
        Next
        
        'sum-up freight of loose cargo
        Dim cFreight As Double
        cFreight = 0
        'txtAmount4.Text = "0.00"
        For i = 1 To .Rows - 1
          cFreight = cFreight + toNumber(.TextMatrix(i, 10))
        Next
        txtAmount4.Text = toMoney(cFreight)
        
        
        'if item is alone
        If Grid.Rows = 2 And Grid.TextMatrix(1, 1) <> "" Then Grid.TextMatrix(1, 10) = 100
        
        'clear boxes
        txtOQty.Text = ""
        txtRQty.Text = ""
        dcUnit.Text = ""
        txtValue.Text = ""
        cboFindList cboClass, ""
        txtGross(0).Text = ""
        txtDisc.Text = ""
        txtNetAmount.Text = ""
        txtFreight.Text = ""
        
'        'Add the amount to current load amount
'        cIGross = cIGross + toNumber(txtGross(0).Text)
'        txtGross(2).Text = Format$(cIGross, "#,##0.00")
'        cIAmount = cIAmount + toNumber(txtNetAmount.Text)
'        cDAmount = cDAmount + toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(txtRQty.Text) * toNumber(txtValue.Text)))
'        txtDesc.Text = Format$(cDAmount, "#,##0.00")
'        txtNet.Text = Format$(cIAmount, "#,##0.00")
'        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
'        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
'
'        txtAmount1_Change
        
        
        
        'Highlight the current row's column
        .ColSel = 10
        'Display a remove button
        'Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
    
    btnUpdate.Enabled = False
End Sub

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update grooss to current purchase amount
        cIGross = cIGross - toNumber(Grid.TextMatrix(.RowSel, 7))
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        'Update amount to current invoice amount
        cIAmount = cIAmount - toNumber(Grid.TextMatrix(.RowSel, 9))
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        'Update discount to current invoice disc
        cDAmount = cDAmount - toNumber(toNumber(txtDisc.Text) / 100) * (toNumber(toNumber(Grid.TextMatrix(.RowSel, 4)) * toNumber(Grid.TextMatrix(.RowSel, 6))))
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

Private Function GetFreightOfLooseCargo(ByVal Supplier As String, ByVal Class As String) As Double
  Dim sql As String
  Dim rstemp As New ADODB.Recordset
  
  sql = "SELECT Cargo_Class.Freight " _
  & "FROM Shipping_Company LEFT JOIN Cargo_Class ON Shipping_Company.ShippingCompanyID = Cargo_Class.ShippingCompanyID " _
  & "WHERE (((Shipping_Company.ShippingCompany)='" & Replace(Supplier, "'", "''") & "') AND " _
  & "((Cargo_Class.Class)='" & Replace(Class, "'", "''") & "'))"
  rstemp.Open sql, CN, adOpenDynamic, adLockOptimistic
  
  If Not rstemp.EOF Then
    GetFreightOfLooseCargo = rstemp!Freight
  Else
    GetFreightOfLooseCargo = 0
  End If
  
  
  rstemp.Close
  Set rstemp = Nothing
End Function

Private Sub cboClass_Click()
  txtFreight.Text = toMoney(GetFreightOfLooseCargo(nsdShippingCo.Text, cboClass.Text))
End Sub

Private Sub CmdReturn_Click()
  Dim RSDetails As New Recordset
  
  RSDetails.CursorLocation = adUseClient
  RSDetails.Open "SELECT * FROM qry_Forwarders_Detail WHERE ForwarderID=" & PK & " AND QtyOnDock > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
  If RSDetails.RecordCount > 0 Then
    With frmPOReceiveLocalAE
      .State = adStateAddMode
      .PK = PK
      .show vbModal
    End With
  Else
    MsgBox "All items are already delivered to VT.", vbInformation
  End If
End Sub

Private Sub dcClass_Click(Area As Integer)
  txtAmount4.Enabled = True
  cboClass.Enabled = True
  cboClass.Clear
  If dcClass.Text = "Loose Cargo" Then
    cboClass.AddItem "Bundle by Cases"
    cboClass.AddItem "Bundle by Bags"
    cboClass.AddItem "Sacks"
    txtFreight.Locked = True
    Grid.TextMatrix(0, 10) = "Freight Amt."
    Labels(7).Caption = "Freight Amt."
    
    'dblLoose = dblLoose + GetFreight(nsdShippingCo.Text, dcClass.Text)
    'txtAmount4.Text = toMoney(dblLoose)
  Else
    'If nsdShippingCo.Text = "" Then Exit Sub
    'txtAmount4.Text = toMoney(GetFreight(nsdShippingCo.Text, dcClass.Text))
    'txtFreight.Locked = False
  End If
End Sub


Private Function GetFreight(ByVal Company As String, ByVal Class As String) As Double
'  Dim sql As String
'  Dim rstemp As New ADODB.Recordset
'
'  sql = "SELECT Cargo_Class.Freight " _
'  & "FROM Shipping_Company INNER JOIN Cargo_Class ON Shipping_Company.ShippingCompanyID = Cargo_Class.ShippingCompanyID " _
'  & "WHERE (((Shipping_Company.ShippingCompany)='" & Replace(Company, "'", "''") & "') AND " _
'  & "((Cargo_Class.Class)='" & Replace(Class, "'", "''") & "'))"
'  rstemp.Open sql, CN, adOpenDynamic, adLockOptimistic
'  If Not rstemp.EOF Then
'    GetFreight = rstemp!freight
'  Else
'    GetFreight = 0
'  End If
'
'
'  rstemp.Close
'  Set rstemp = Nothing
End Function

Private Sub nsdLocal_Change()
  Dim sql As String
  Dim rstemp As New ADODB.Recordset
  
  sql = "SELECT Local_Forwarder.LocalForwarderID, Local_Forwarder_Account_Description.AccTitle, Local_Forwarder_Detail.Amount " _
  & "FROM Local_Forwarder_Account_Description RIGHT JOIN (Local_Forwarder LEFT JOIN Local_Forwarder_Detail ON Local_Forwarder.LocalForwarderID = Local_Forwarder_Detail.LocalForwarderID) ON Local_Forwarder_Account_Description.LocalForwarderAccTitleID = Local_Forwarder_Detail.AccountDescriptionID " _
  & "WHERE (((Local_Forwarder.LocalForwarder)='" & Replace(nsdLocal.Text, "'", "''") & "'))"
  
  rstemp.Open sql, CN, adOpenDynamic, adLockOptimistic
  
  txtAmount6.Text = "0.00"
  txtAmount7.Text = "0.00"
  
  Do While Not rstemp.EOF
    If rstemp!AccTitle = "Local Trucking" Then txtAmount6.Text = toMoney(rstemp!Amount)
    If rstemp!AccTitle = "Sidewalk Handling" Then txtAmount7.Text = toMoney(rstemp!Amount)
    
    rstemp.MoveNext
  Loop
  
  rstemp.Close
  Set rstemp = Nothing
End Sub

Private Sub txtAmount1_Change()
  txtTotalTranspoCost.Text = toMoney(toNumber(txtAmount1.Text) + toNumber(txtAmount2.Text) _
  + toNumber(txtAmount3.Text) + toNumber(txtAmount4.Text) + toNumber(txtAmount5.Text) _
  + toNumber(txtAmount6.Text) + toNumber(txtAmount7.Text))
End Sub

Private Sub txtAmount2_Change()
txtAmount1_Change
End Sub

Private Sub txtAmount3_Change()
txtAmount1_Change
End Sub

Private Sub txtAmount4_Change()
'txtAmount1_Change
End Sub

Private Sub txtAmount5_Change()
txtAmount1_Change
End Sub

Private Sub txtAmount6_Change()
txtAmount1_Change
End Sub

Private Sub txtAmount7_Change()
txtAmount1_Change
End Sub

Private Sub txtCostPerPackage_Click()
'  MsgBox Grid.Row
End Sub

Private Sub txtdisc_Change()
  If Trim(txtDisc.Text) = "" Then txtDisc.Text = 0
  
  txtRQty_Change
End Sub

Private Sub txtdisc_Click()
    txtQty_Change
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub txtDisc_GotFocus()
    HLText txtDisc
End Sub

Private Sub txtdisc_Validate(Cancel As Boolean)
    txtDisc.Text = toNumber(txtDisc.Text)
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewerPH.INV_PK = PK
    'frmInvoiceViewerPH.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.lblTitle.Caption = "Payment History Viewer"
    'frmInvoiceViewerPH.show vbModal
End Sub

Private Sub cmdSave_Click()
    'Verify the entries
    
    If Trim(nsdShippingCo.Text) = "" Then
        MsgBox "Please enter shipping company before saving.", vbExclamation
        Exit Sub
    End If
    If (dcClass.Text = "A" Or dcClass.Text = "B" Or dcClass.Text = "C") And Trim(txtVanNo.Text) = "" Then
       MsgBox "Please enter Van No. before saving.", vbExclamation
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to return before saving this record.", vbExclamation
        Exit Sub
    End If
    
    'check if freight allocation is 100 percent
    Dim i As Integer
    Dim j As Double
    j = 0
    For i = 1 To Grid.Rows - 1
      j = j + toNumber(Grid.TextMatrix(i, 10))
    Next
    
    If (Grid.Rows > 2) And (j <> 100) Then
      MsgBox "System detects that your freight allocation has a problem, please make correction before saving.", vbExclamation
      Exit Sub
    End If
    '----
    
    
    
    
   
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    'Connection for Forwarders
    Dim RSShipping As New Recordset

    RSShipping.CursorLocation = adUseClient
    RSShipping.Open "Forwarders", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    'Connection for Forwarders_Detail
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "Forwarders_Detail", CN, adOpenDynamic, adLockOptimistic, adCmdTable

    'Connection for Transportation_Cost
    Dim RSTransport As New Recordset

    RSTransport.CursorLocation = adUseClient
    RSTransport.Open "Transportation_Cost", CN, adOpenDynamic, adLockOptimistic, adCmdTable


    Screen.MousePointer = vbHourglass

    Dim c As Integer

    On Error GoTo err

    CN.BeginTrans

    'Save the record
    With RSShipping
        .AddNew
        Dim ShippingPK As Integer
        
        ShippingPK = getIndex("Forwarders")
        ![POID] = PK
        ![DRNo] = ShippingPK
        ![ForwarderID] = ShippingPK
        ![ShippingCompanyID] = nsdShippingCo.BoundText
        ![ShippingGuideNo] = txtShippingGuideNo.Text
        ![Ship] = txtShip.Text
        ![ArrivalDate] = dtpShippingDate.Value
        ![DRNo] = txtDRNo.Text
        ![BLNo] = txtBLNo.Text
        ![TruckNo] = txtTruckNo.Text
        ![VanNo] = txtVanNo.Text
        ![VoyageNo] = txtVoyageNo.Text
        ![Ambot1] = cboAmbot.Text
        
        ![ReceiptNo] = txtReceipt.Text
        ![ReceiptDate] = dtpReceiptDate.Value
        ![CheckNo] = txtCheckNo.Text
        ![SendThru] = cboSendThru.Text
        ![Total] = CDbl(txtTotal.Text)
        ![CostPerPackage] = toNumber(txtCostPerPackage.Text)
        
        ![Gross] = toNumber(txtGross(2).Text)
        ![Discount] = txtDesc.Text
        ![TaxBase] = toNumber(txtTaxBase.Text)
        ![Vat] = toNumber(txtVat.Text)
        ![NetAmount] = toNumber(txtNet.Text)
        ![LocalForwarderID] = nsdLocal.BoundText
        
        ![DateAdded] = Now
        ![AddedByFK] = CurrUser.USER_PK
                
        .Update
    End With
   
   
   'Save the record
    With RSTransport
        .AddNew
                
        ![ForwarderID] = ShippingPK
        !MlaTruckingDate = dtp1.Value: !MlaTruckingOR = txtOR1.Text: !MlaTruckingAmount = txtAmount1.Text
        !MlaArrastreDate = dtp1.Value: !MlaArrastreOR = txtOR2.Text: !MlaArrastreAmount = txtAmount2.Text
        !MlaWfgFeeDate = dtp1.Value: !MlaWfgFeeOR = txtOR3.Text: !MlaWfgFeeAmount = txtAmount3.Text
        !FreightDate = dtp1.Value: !FreightOR = txtOR4.Text: !FreightAmount = txtAmount4.Text
        !LocalArrastreDate = dtp1.Value: !LocalArrastreOR = txtOR5.Text: !LocalArrastreAmount = txtAmount5.Text
        !LocalTruckingDate = dtp1.Value: !LocalTruckingOR = txtOR6.Text: !LocalTruckingAmount = txtAmount6.Text
        !SidewalkHandlingDate = dtp1.Value: !SidewalkHandlingOR = txtOR7.Text: !SidewalkHandlingAmount = txtAmount7.Text
                
        .Update
    End With
   
   
   
   
   
    With Grid
        
        'Save to Shipping Guide Details
        Dim RSSPurchaseOrderDetails As New Recordset
    
        RSSPurchaseOrderDetails.CursorLocation = adUseClient
        RSSPurchaseOrderDetails.Open "SELECT * From Purchase_Order_Detail where POID = " & PK, CN, , adLockOptimistic
        
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            
            RSDetails.AddNew

            RSDetails![ForwarderID] = ShippingPK
            RSDetails![StockID] = toNumber(.TextMatrix(c, 11))
            RSDetails![Qty] = toNumber(.TextMatrix(c, 3))
            RSDetails![UnitID] = getUnitID(.TextMatrix(c, 4))
            RSDetails![Value] = CDbl(.TextMatrix(c, 5))
            RSDetails![Class] = .TextMatrix(c, 6)
            RSDetails![FreightPercent] = toNumber(.TextMatrix(c, 10))

            RSDetails.Update

            

            
            
            'add qty received in Purchase Order Details
            RSSPurchaseOrderDetails.Find "[StockID] = " & toNumber(.TextMatrix(c, 11)), , adSearchForward, 1
            RSSPurchaseOrderDetails!QtyReceived = toNumber(RSSPurchaseOrderDetails!QtyReceived) + toNumber(.TextMatrix(c, 3))
            
            RSSPurchaseOrderDetails.Update
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
    If CloseMe = True Then
        Unload Me
    Else
        nsdShippingCo.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
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


Private Function ShippingGuideNo() As Long
    ShippingGuideNo = getIndex("Forwarders")
End Function


Private Sub Form_Load()
  dblLoose = 0
  
  cCostPerPackage = 0
  cTotalAmount = 0
  cTotalTranspoCost = 0
  cIGross = 0
  cIAmount = 0
  
  
  
  InitGrid

    bind_dc "SELECT * FROM Unit", "Unit", dcUnit, "UnitID", True

    Screen.MousePointer = vbHourglass
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
      InitNSD
      
      'Set the recordset
      If rs.State = 1 Then rs.Close
      rs.Open "SELECT * FROM qry_Purchase_Order WHERE POID=" & PK, CN, adOpenStatic, adLockOptimistic
      dtpShippingDate.Value = Date
      dtpReceiptDate.Value = Date
      Caption = "Create New Entry"
      cmdUsrHistory.Enabled = False
      
      
      txtShippingGuideNo.Text = Format(Date, "yy") & Format(Date, "mm") & Format(ShippingGuideNo, "0000")
      DisplayForEditing
    Else
        'Set the recordset
        If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM qry_Forwarders1 WHERE ForwarderID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        cmdCancel.Caption = "Close"
        DisplayForViewing
        
        If ForCusAcc = True Then
            'Me.Icon = frmLocalPurchaseReturn.Icon
        End If
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
    PK = getIndex("Forwarders")
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
        .ColSel = 10
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 2775
        .ColWidth(2) = 510
        .ColWidth(3) = 578
        .ColWidth(4) = 975
        .ColWidth(5) = 1110
        .ColWidth(6) = 1440
        .ColWidth(7) = 1110
        .ColWidth(8) = 615
        .ColWidth(9) = 1110
        .ColWidth(10) = 800
        .ColWidth(11) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product"
        .TextMatrix(0, 2) = "O-Qty"
        .TextMatrix(0, 3) = "R-Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Price"
        .TextMatrix(0, 6) = "Class"
        .TextMatrix(0, 7) = "Gross"
        .TextMatrix(0, 8) = "Disc %"
        .TextMatrix(0, 9) = "Net Amount"
        .TextMatrix(0, 10) = "Freight %"
        .TextMatrix(0, 11) = "Stock ID"
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

Private Sub ResetEntry()
    'nsdStock.ResetValue
    'txtUnitPrice.Tag = 0
    txtValue.Text = "0.00"
    txtOQty.Text = 0
    txtRQty.Text = 0
    txtStock.Text = ""
    dcUnit.Text = ""
    cboFindList cboClass, ""
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'If HaveAction = True Then
    '    frmLocalPurchaseReturn.RefreshRecords
    'End If
    
    Set frmPOReceiveLocalAE = Nothing
End Sub

Private Sub Grid_Click()
 ' MsgBox Grid.ColWidth(1) & vbCr _
  & Grid.ColWidth(2) & vbCr _
  & Grid.ColWidth(3) & vbCr _
  & Grid.ColWidth(4) & vbCr _
  & Grid.ColWidth(5) & vbCr _
  & Grid.ColWidth(6) & vbCr _
  & Grid.ColWidth(7) & vbCr _
  & Grid.ColWidth(8) & vbCr _
  & Grid.ColWidth(9) & vbCr _
  & Grid.ColWidth(10)
  
  'Exit Sub
  btnUpdate.Enabled = True
  
    With Grid
        txtStock.Text = .TextMatrix(.RowSel, 1)
        txtOQty.Text = toNumber(.TextMatrix(.RowSel, 2))
        txtRQty = .TextMatrix(.RowSel, 3)
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        txtValue.Text = toMoney(.TextMatrix(.RowSel, 5))
        cboFindList cboClass, .TextMatrix(.RowSel, 6)
        txtGross(0).Text = toMoney(.TextMatrix(.RowSel, 7))
        txtDisc.Text = toNumber(.TextMatrix(.RowSel, 8))
        txtNetAmount.Text = toMoney(.TextMatrix(.RowSel, 9))
        txtFreight.Text = toNumber(.TextMatrix(.RowSel, 10))
        
        
        If State = adStateEditMode Then Exit Sub
        If Grid.Rows = 2 And Grid.TextMatrix(1, 10) = "" Then
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
        
        
        
        'compute cost per package
        If dcClass.Text = "A" Or dcClass.Text = "B" Or dcClass.Text = "C" Then
          If Grid.TextMatrix(Grid.Row, 3) = "0" Then
            txtCostPerPackage.Text = "0.00"
          Else
            txtCostPerPackage.Text = toMoney(((toNumber(txtAmount1.Text) + toNumber(txtAmount2.Text) + _
            toNumber(txtAmount3.Text) + toNumber(txtAmount4.Text) + toNumber(txtAmount5.Text)) / _
            toNumber(Grid.TextMatrix(Grid.Row, 3))) + toNumber(txtAmount6.Text) + toNumber(txtAmount7.Text))
          End If
        Else
          If Grid.TextMatrix(Grid.Row, 3) = "0" Or Grid.TextMatrix(Grid.Row, 10) = "0" Then
            txtCostPerPackage.Text = "0.00"
          Else
            txtCostPerPackage.Text = toMoney((((toNumber(txtAmount1.Text) + toNumber(txtAmount2.Text) + _
            toNumber(txtAmount3.Text) + toNumber(txtAmount5.Text)) * toNumber((Grid.TextMatrix(Grid.Row, 10)))) / _
            toNumber(Grid.TextMatrix(Grid.Row, 3))) + toNumber(txtAmount4.Text) + toNumber(txtAmount6.Text) + toNumber(txtAmount7.Text))
          End If
        End If
        
    End With
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub

Private Sub nsdShippingCo_Change()
  bind_dc "SELECT * FROM qry_Cargo_Class where shippingcompany='" _
  & Replace(nsdShippingCo.Text, "'", "''") & "'", "Cargo", dcClass, "CargoClassID", True
  dcClass.Text = ""
  
End Sub

Private Sub txtCostPerPackage_GotFocus()
  HLText txtCostPerPackage
End Sub

Private Sub txtCostPerPackage_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtCostPerPackage)
End Sub

Private Sub txtDate_GotFocus()
    HLText txtDate
End Sub

Private Sub txtDesc_GotFocus()
    HLText txtDesc
End Sub

'Private Sub txtQty_LostFocus()
    'If txtQty > intQtyOld Then
    '    MsgBox "Overdelivery for " & txtStock.Text & ".", vbInformation
    '    txtQty.Text = intQtyOld
    'End If
'End Sub

'Private Sub txtQty_Validate(Cancel As Boolean)
'    txtQty.Text = toNumber(txtQty.Text)
'End Sub

Private Sub txtUnitPrice_Change()
    txtQty_Change
End Sub

'Private Sub txtUnitPrice_Validate(Cancel As Boolean)
'    txtUnitPrice.Text = toMoney(toNumber(txtUnitPrice.Text))
'End Sub

Private Sub txtQty_Change()
    If toNumber(txtRQty.Text) < 1 Then
        btnUpdate.Enabled = False
    Else
        btnUpdate.Enabled = True
    End If
    
    txtGross(0).Text = toMoney((toNumber(txtRQty.Text) * toNumber(txtValue.Text)))
    txtNetAmount.Text = toMoney((toNumber(txtRQty.Text) * toNumber(txtValue.Text)) - ((toNumber(txtDisc.Text) / 100) * toNumber(toNumber(txtRQty.Text) * toNumber(txtValue.Text))))
End Sub

Private Sub txtUnitPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

'Used to edit record
Private Sub DisplayForEditing()
    On Error GoTo err
    
    txtSupplier.Text = rs!Company
    txtPONo.Text = rs!PONo
    txtAddress.Text = rs!Address
    txtDate.Text = rs![Date]
        
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Purchase_Order_Detail WHERE POID=" & PK & " AND QtyDue > 0 ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        
        
        With Grid
          .Rows = 1
          While Not RSDetails.EOF
            .Rows = .Rows + 1
            If .Rows = 2 And .TextMatrix(1, 10) = "" Then
              .TextMatrix(1, 1) = RSDetails![Stock]
              .TextMatrix(1, 2) = IIf(RSDetails![QtyDue] = 0, 0, RSDetails![QtyDue])
              .TextMatrix(1, 3) = "0"
              .TextMatrix(1, 4) = RSDetails![unit]
              .TextMatrix(1, 5) = toMoney(RSDetails![Price])
              .TextMatrix(1, 11) = RSDetails![StockID]
            Else
              
              .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
              .TextMatrix(.Rows - 1, 2) = IIf(RSDetails![QtyReceived] = 0, RSDetails![QtyDue], RSDetails![QtyReceived])
              .TextMatrix(.Rows - 1, 3) = "0"
              .TextMatrix(.Rows - 1, 4) = RSDetails![unit]
              .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Price])
              .TextMatrix(.Rows - 1, 11) = RSDetails![StockID]
            End If
            cIRowCount = cIRowCount + 1
            
          
            RSDetails.MoveNext
          Wend
          
          
          'dont ask
          If .Rows = 1 Then .Rows = 2
          .FixedRows = 1
          
        End With
        
        
        
        
        Grid.Row = 1
        Grid.ColSel = 10
        'Set fixed cols
        If State = adStateEditMode Then
          Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
          Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
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
    txtAddress.Enabled = True
    txtDate.Enabled = True
    
    txtSupplier.Text = rs!Company
    txtPONo.Text = rs!PONo
    txtAddress.Text = rs!Address
    txtDate.Text = rs![Date]
    
    nsdShippingCo.Text = rs![ShippingCompany]
    txtDRNo.Text = rs![DRNo]
    txtShippingGuideNo.Text = rs![ShippingGuideNo]
    dtpShippingDate.Value = rs![ArrivalDate]
    txtBLNo.Text = rs![BLNo]
    txtVoyageNo.Text = rs![VoyageNo]
    txtShip.Text = rs![Ship]
    txtVanNo.Text = rs![VanNo]
    txtTruckNo.Text = rs![TruckNo]
    cboFindList cboAmbot, rs![Ambot1]
    txtReceipt.Text = rs![ReceiptNo]
    dtpReceiptDate.Value = rs![ReceiptDate]
    txtCheckNo.Text = rs![CheckNo]
    cboFindList cboSendThru, rs![SendThru]
    nsdLocal.Text = rs![LocalForwarder]
    dcClass.Text = rs![Class]
    
    txtTotal.Text = toMoney(toNumber(rs![Total]))
    txtCostPerPackage.Text = toMoney(toNumber(rs![CostPerPackage]))
    
    
    
    'Display the Transport Cost
    Dim RSTransport As New Recordset
    RSTransport.CursorLocation = adUseClient
    RSTransport.Open "SELECT * FROM qry_Transport_Cost WHERE ForwarderID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    With RSTransport
      dtp1.Value = IIf(IsNull(!MlaTruckingDate), "", !MlaTruckingDate): txtOR1.Text = IIf(IsNull(!MlaTruckingOR), "", !MlaTruckingOR): txtAmount1.Text = toMoney(!MlaTruckingAmount)
      dtp2.Value = IIf(IsNull(!MlaArrastreDate), "", !MlaArrastreDate): txtOR2.Text = IIf(IsNull(!MlaArrastreOR), "", !MlaArrastreOR): txtAmount2.Text = toMoney(!MlaArrastreAmount)
      dtp3.Value = IIf(IsNull(!MlaWfgFeeDate), "", !MlaWfgFeeDate): txtOR3.Text = IIf(IsNull(!MlaWfgFeeOR), "", !MlaWfgFeeOR): txtAmount3.Text = toMoney(!MlaWfgFeeAmount)
      dtp4.Value = IIf(IsNull(!FreightDate), "", !FreightDate): txtOR4.Text = IIf(IsNull(!FreightOR), "", !FreightOR): txtAmount4.Text = toMoney(!FreightAmount)
      dtp5.Value = IIf(IsNull(!LocalArrastreDate), "", !LocalArrastreDate): txtOR5.Text = IIf(IsNull(!LocalArrastreOR), "", !LocalArrastreOR): txtAmount5.Text = toMoney(!LocalArrastreAmount)
      dtp6.Value = IIf(IsNull(!LocalTruckingDate), "", !LocalTruckingDate): txtOR6.Text = IIf(IsNull(!LocalTruckingOR), "", !LocalTruckingOR): txtAmount6.Text = toMoney(!LocalTruckingAmount)
      dtp7.Value = IIf(IsNull(!SidewalkHandlingDate), "", !SidewalkHandlingDate): txtOR7.Text = IIf(IsNull(!SidewalkHandlingOR), "", !SidewalkHandlingOR): txtAmount7.Text = toMoney(!SidewalkHandlingAmount)
    
    End With
    
    
    
    
    
    
    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Forwarders1_Detail WHERE ForwarderID=" & PK & " ORDER BY Stock ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 10) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Stock]
                    .TextMatrix(1, 2) = RSDetails![OQty]
                    .TextMatrix(1, 3) = RSDetails![RQty]
                    .TextMatrix(1, 4) = RSDetails![unit]
                    .TextMatrix(1, 5) = toMoney(RSDetails![Value])
                    .TextMatrix(1, 6) = RSDetails![Class]
                    .TextMatrix(1, 7) = RSDetails![Gross]
                    .TextMatrix(1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(1, 10) = RSDetails![FreightPercent]
                    .TextMatrix(1, 11) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![OQty]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![RQty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![unit]
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSDetails![Value])
                    .TextMatrix(.Rows - 1, 6) = RSDetails![Class]
                    .TextMatrix(.Rows - 1, 7) = RSDetails![Gross]
                    .TextMatrix(.Rows - 1, 8) = RSDetails![Discount] * 100
                    .TextMatrix(.Rows - 1, 9) = toMoney(RSDetails![NetAmount])
                    .TextMatrix(.Rows - 1, 10) = RSDetails![StockID]
                End If
            End With
            
            cIGross = cIGross + toNumber(RSDetails![Gross])
            
            cIAmount = cIAmount + toNumber(RSDetails![NetAmount])
            cDAmount = cDAmount + toNumber(toNumber(RSDetails![Discount]) / 100) * (toNumber(toNumber(RSDetails![RQty]) * toNumber(RSDetails![Value])))
            
            
            
            RSDetails.MoveNext
        Wend
        txtGross(2).Text = Format$(cIGross, "#,##0.00")
        txtDesc.Text = Format$(cDAmount, "#,##0.00")
        txtNet.Text = Format$(cIAmount, "#,##0.00")
        txtTaxBase.Text = toMoney(txtNet.Text / 1.12)
        txtVat.Text = toMoney(txtNet.Text - txtTaxBase.Text)
        
        
        
        Grid.Row = 1
        Grid.ColSel = 10
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            'Grid.FixedCols = 2
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
  
    'Disable commands
    LockInput Me, True

    'dtpInvoiceDate.Visible = True
    'txtInvoiceDate.Visible = False
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
    Grid.Top = 3090
    Grid.Height = 2970
    
    Exit Sub
err:
    'Error if encounter a null value
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Description
    End If
End Sub

'Private Sub txtUnitPrice_GotFocus()
'    HLText txtUnitPrice
'End Sub





Private Sub txtFreight_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtFreight)
End Sub

Private Sub txtOQty_GotFocus()
  HLText txtOQty
End Sub

Private Sub txtOQty_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtOQty)
End Sub

Private Sub txtRQty_Change()
  If Trim(txtRQty.Text) = "" Then txtRQty.Text = 0
  
  txtGross(0).Text = toMoney(toNumber(txtRQty.Text) * toNumber(txtValue.Text))
  txtNetAmount.Text = toMoney(txtGross(0).Text - (txtGross(0).Text * (toNumber(txtDisc.Text) / 100)))
End Sub

Private Sub txtRQty_GotFocus()
  HLText txtRQty
End Sub

Private Sub txtRQty_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtRQty)
End Sub

Private Sub txtTotal_GotFocus()
  HLText txtTotal
End Sub

Private Sub txtTotal_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtTotal)
End Sub

Private Sub txtValue_Change()
    If Trim(txtValue.Text) = "" Then txtValue.Text = 0
    
    txtRQty_Change
End Sub

Private Sub txtValue_GotFocus()
  HLText txtValue
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
  KeyAscii = AllowOnlyNumbers(KeyAscii, txtValue)
End Sub
