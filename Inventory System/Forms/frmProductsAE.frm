VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmProductsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   7410
   ClientLeft      =   1365
   ClientTop       =   2085
   ClientWidth     =   10275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10275
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   6
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   6
      Top             =   2040
      Width           =   1320
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3495
      Left            =   210
      TabIndex        =   30
      Top             =   3210
      Width           =   9885
      _ExtentX        =   17436
      _ExtentY        =   6165
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Product Measures"
      TabPicture(0)   =   "frmProductsAE.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label7"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label8"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "nsdUnit"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Grid"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtSalesPrice"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "btnRemove"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtOrder"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtOnHand"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdAdd"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtQty"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtSupplierPrice"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtPending"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtIncoming"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cmdAdjust"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).ControlCount=   20
      TabCaption(1)   =   "Price History"
      TabPicture(1)   =   "frmProductsAE.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lvPriceHistory"
      Tab(1).ControlCount=   1
      Begin VB.CommandButton cmdAdjust 
         Caption         =   "Adjust"
         Height          =   315
         Left            =   8100
         TabIndex        =   43
         Top             =   690
         Width           =   615
      End
      Begin VB.TextBox txtIncoming 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5910
         MaxLength       =   10
         TabIndex        =   15
         Text            =   "0"
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox txtPending 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5130
         MaxLength       =   10
         TabIndex        =   14
         Text            =   "0"
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox txtSupplierPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3870
         MaxLength       =   10
         TabIndex        =   13
         Text            =   "0.00"
         Top             =   690
         Width           =   1200
      End
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   780
         MaxLength       =   10
         TabIndex        =   10
         Top             =   690
         Width           =   540
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   7500
         TabIndex        =   17
         Top             =   690
         Width           =   495
      End
      Begin VB.TextBox txtOnHand 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6720
         MaxLength       =   10
         TabIndex        =   16
         Text            =   "0"
         Top             =   690
         Width           =   720
      End
      Begin VB.TextBox txtOrder 
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Top             =   690
         Width           =   585
      End
      Begin VB.CommandButton btnRemove 
         Height          =   275
         Left            =   210
         Picture         =   "frmProductsAE.frx":0038
         Style           =   1  'Graphical
         TabIndex        =   32
         ToolTipText     =   "Remove"
         Top             =   1200
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.TextBox txtSalesPrice 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2850
         MaxLength       =   10
         TabIndex        =   12
         Text            =   "0.00"
         Top             =   690
         Width           =   960
      End
      Begin MSComctlLib.ListView lvPriceHistory 
         Height          =   2895
         Left            =   -74880
         TabIndex        =   31
         Top             =   450
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   5106
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   2196
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Supplier"
            Object.Width           =   10769
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Price"
            Object.Width           =   2540
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   2220
         Left            =   150
         TabIndex        =   33
         Top             =   1080
         Width           =   8565
         _ExtentX        =   15108
         _ExtentY        =   3916
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
      Begin ctrlNSDataCombo.NSDataCombo nsdUnit 
         Height          =   315
         Left            =   1380
         TabIndex        =   11
         Top             =   690
         Width           =   1440
         _ExtentX        =   2540
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
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Incoming"
         Height          =   255
         Left            =   5910
         TabIndex        =   41
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Pending"
         Height          =   255
         Left            =   5130
         TabIndex        =   40
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Supplier Price"
         Height          =   255
         Left            =   3870
         TabIndex        =   39
         Top             =   420
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty"
         Height          =   285
         Left            =   810
         TabIndex        =   38
         Top             =   420
         Width           =   435
      End
      Begin VB.Label Label2 
         Caption         =   "Unit"
         Height          =   285
         Left            =   1410
         TabIndex        =   37
         Top             =   420
         Width           =   705
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "On-Hand"
         Height          =   255
         Left            =   6690
         TabIndex        =   36
         Top             =   420
         Width           =   675
      End
      Begin VB.Label Label5 
         Caption         =   "Order"
         Height          =   255
         Left            =   180
         TabIndex        =   35
         Top             =   420
         Width           =   555
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Price"
         Height          =   255
         Left            =   2850
         TabIndex        =   34
         Top             =   420
         Width           =   945
      End
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmProductsAE.frx":01EA
      Left            =   1500
      List            =   "frmProductsAE.frx":01F7
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   2760
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   210
      TabIndex        =   20
      Top             =   6930
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   5
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   4
      Top             =   1680
      Width           =   780
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1335
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   2
      Top             =   990
      Width           =   4125
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   630
      Width           =   7365
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8730
      TabIndex        =   19
      Top             =   6960
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   7320
      TabIndex        =   18
      Top             =   6960
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   1500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   1965
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   240
      TabIndex        =   21
      Top             =   6825
      Width           =   9795
      _ExtentX        =   17277
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1500
      TabIndex        =   7
      Top             =   2400
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcReoderUnit 
      Height          =   315
      Left            =   2790
      TabIndex        =   5
      Top             =   1680
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
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
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Ext Price"
      Height          =   240
      Index           =   3
      Left            =   210
      TabIndex        =   42
      Top             =   2040
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Unit"
      Height          =   285
      Left            =   2370
      TabIndex        =   29
      Top             =   1680
      Width           =   345
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   240
      Index           =   9
      Left            =   210
      TabIndex        =   28
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   210
      TabIndex        =   27
      Top             =   2370
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Reorder Pt."
      Height          =   240
      Index           =   7
      Left            =   210
      TabIndex        =   26
      Top             =   1680
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "ICode"
      Height          =   240
      Index           =   4
      Left            =   210
      TabIndex        =   25
      Top             =   1365
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Short"
      Height          =   240
      Index           =   2
      Left            =   210
      TabIndex        =   24
      Top             =   945
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Product"
      Height          =   240
      Index           =   1
      Left            =   210
      TabIndex        =   23
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Barcode"
      Height          =   240
      Index           =   0
      Left            =   210
      TabIndex        =   22
      Top             =   285
      Width           =   1215
   End
End
Attribute VB_Name = "frmProductsAE"
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

Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim RS                      As New Recordset
Dim rs1                     As New Recordset
Dim RSStockUnit             As New Recordset

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub cmdAdd_Click()
    If Trim(txtOrder.Text) = "" Or Trim(txtQty.Text) = "" Or Trim(nsdUnit.Text) = "" Then Exit Sub

    Dim CurrRow As Integer
    Dim intUnitID As Integer
    
'    If nsdUnit.BoundText = "" Then
        CurrRow = getFlexPos(Grid, 9, nsdUnit.Tag)
'        intUnitID = nsdUnit.Tag
'    Else
'        CurrRow = getFlexPos(Grid, 9, nsdUnit.BoundText)
'        intUnitID = nsdUnit.BoundText
'    End If

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 9) = "" Then
                .TextMatrix(1, 1) = txtOrder.Text
                .TextMatrix(1, 2) = txtQty.Text
                .TextMatrix(1, 3) = nsdUnit.Text
                .TextMatrix(1, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(1, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(1, 6) = txtPending.Text
                .TextMatrix(1, 7) = txtIncoming.Text
                .TextMatrix(1, 8) = txtOnHand.Text
                .TextMatrix(1, 9) = nsdUnit.Tag 'intUnitID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = txtOrder.Text
                .TextMatrix(.Rows - 1, 2) = txtQty.Text
                .TextMatrix(.Rows - 1, 3) = nsdUnit.Text
                .TextMatrix(.Rows - 1, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(.Rows - 1, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(.Rows - 1, 6) = txtPending.Text
                .TextMatrix(.Rows - 1, 7) = txtIncoming.Text
                .TextMatrix(.Rows - 1, 8) = txtOnHand.Text
                .TextMatrix(.Rows - 1, 9) = nsdUnit.Tag 'intUnitID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = txtOrder.Text
                .TextMatrix(CurrRow, 2) = txtQty.Text
                .TextMatrix(CurrRow, 3) = nsdUnit.Text
                .TextMatrix(CurrRow, 4) = toMoney(txtSalesPrice.Text)
                .TextMatrix(CurrRow, 5) = toMoney(txtSupplierPrice.Text)
                .TextMatrix(CurrRow, 6) = txtPending.Text
                .TextMatrix(CurrRow, 7) = txtIncoming.Text
                .TextMatrix(CurrRow, 8) = txtOnHand.Text
                .TextMatrix(CurrRow, 9) = nsdUnit.Tag 'intUnitID
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .ColSel = 8
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With RS
      txtEntry(1).Text = .Fields("Barcode")
      txtEntry(2).Text = .Fields("Stock")
      txtEntry(3).Text = .Fields("Short")
      txtEntry(4).Text = .Fields("ICode")
      txtEntry(5).Text = .Fields("ReorderPoint")
      txtEntry(6).Text = toMoney(.Fields("ExtPrice"))
      dcReoderUnit.BoundText = IIf(IsNull(.Fields![UnitID]), "", .Fields![UnitID])
      cboStatus.Text = .Fields("Status")
      dcCategory.BoundText = .Fields![CategoryID]
    End With
    
    'Display the details
    Dim RSStockUnit As New Recordset

    cIRowCount = 0
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID=" & PK & " Order by [Order] ASC", CN, adOpenStatic, adLockOptimistic
    
    If RSStockUnit.RecordCount > 0 Then
        RSStockUnit.MoveFirst
        While Not RSStockUnit.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 9) = "" Then
                    .TextMatrix(1, 1) = RSStockUnit![Order]
                    .TextMatrix(1, 2) = RSStockUnit![Qty]
                    .TextMatrix(1, 3) = RSStockUnit![Unit]
                    .TextMatrix(1, 4) = toMoney(RSStockUnit![SalesPrice])
                    .TextMatrix(1, 5) = toMoney(RSStockUnit![SupplierPrice])
                    .TextMatrix(1, 6) = RSStockUnit![Pending]
                    .TextMatrix(1, 7) = RSStockUnit![Incoming]
                    .TextMatrix(1, 8) = RSStockUnit![Onhand]
                    .TextMatrix(1, 9) = RSStockUnit![UnitID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSStockUnit![Order]
                    .TextMatrix(.Rows - 1, 2) = RSStockUnit![Qty]
                    .TextMatrix(.Rows - 1, 3) = RSStockUnit![Unit]
                    .TextMatrix(.Rows - 1, 4) = toMoney(RSStockUnit![SalesPrice])
                    .TextMatrix(.Rows - 1, 5) = toMoney(RSStockUnit![SupplierPrice])
                    .TextMatrix(.Rows - 1, 6) = RSStockUnit![Pending]
                    .TextMatrix(.Rows - 1, 7) = RSStockUnit![Incoming]
                    .TextMatrix(.Rows - 1, 8) = RSStockUnit![Onhand]
                    .TextMatrix(.Rows - 1, 9) = RSStockUnit![UnitID]
                End If
            End With
            RSStockUnit.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 8
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSStockUnit.Close
    'Clear variables
    Set RSStockUnit = Nothing
    
    Exit Sub
err:
        'If err.Number = 94 Then Resume Next
        MsgBox "Error: " & err.Description, vbExclamation
End Sub

Private Sub cmdAdjust_Click()
    With frmTransferQty
        .StockID = PK
        
        .show 1
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    txtEntry(1).SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    'check for blank product
'    If is_empty(txtEntry(2).Text) = True Then
'        MsgBox "Product should not be empty.", vbExclamation
'        Exit Sub
'    End If
    
    If txtEntry(2).Text = "" Then
        MsgBox "Product should not be empty.", vbExclamation
        Exit Sub
    End If
    
    'check for blank category
    If Trim(dcCategory.Text) = "" Then
        MsgBox "Category should not be empty.", vbExclamation
        Exit Sub
    End If
    
    'check for blank unit measures
    If cIRowCount < 1 Then
        MsgBox "Please provide at least one product measure.", vbExclamation
        Exit Sub
    End If
    
    CN.BeginTrans
      
    If State = adStateAddMode Or State = adStatePopupMode Then
        RS.AddNew
        RS.Fields("StockId") = PK
        RS.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        RS.Fields("datemodified") = Now
        RS.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With RS
        .Fields("Barcode") = txtEntry(1).Text
        .Fields("Stock") = txtEntry(2).Text
        .Fields("Short") = txtEntry(3).Text
        .Fields("ICode") = txtEntry(4).Text
        .Fields("ReorderPoint") = toNumber(txtEntry(5).Text)
        .Fields("ExtPrice") = toMoney(txtEntry(6).Text)
        .Fields("UnitID") = dcReoderUnit.BoundText
        .Fields("Status") = cboStatus.Text
        .Fields("CategoryID") = dcCategory.BoundText
        
        .Update
    End With
  
    Dim RSStockUnit As New Recordset

    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM Stock_Unit WHERE StockID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSStockUnit.AddNew

                RSStockUnit![StockID] = PK
                RSStockUnit![Order] = toNumber(.TextMatrix(c, 1))
                RSStockUnit![UnitID] = toNumber(.TextMatrix(c, 9))
                RSStockUnit![Qty] = toNumber(.TextMatrix(c, 2))
                RSStockUnit![SalesPrice] = toNumber(.TextMatrix(c, 4))
                RSStockUnit![SupplierPrice] = toNumber(.TextMatrix(c, 5))
                RSStockUnit![Pending] = toNumber(.TextMatrix(c, 6))
                RSStockUnit![Incoming] = toNumber(.TextMatrix(c, 7))
                RSStockUnit![Onhand] = toNumber(.TextMatrix(c, 8))

                RSStockUnit.Update
            ElseIf State = adStateEditMode Then
                RSStockUnit.Filter = "UnitID = " & toNumber(.TextMatrix(c, 9))
            
                If RSStockUnit.RecordCount = 0 Then GoTo AddNew

                RSStockUnit![Order] = toNumber(.TextMatrix(c, 1))
                RSStockUnit![UnitID] = toNumber(.TextMatrix(c, 9))
                RSStockUnit![Qty] = toNumber(.TextMatrix(c, 2))
                RSStockUnit![SalesPrice] = toNumber(.TextMatrix(c, 4))
                RSStockUnit![SupplierPrice] = toNumber(.TextMatrix(c, 5))
                RSStockUnit![Pending] = toNumber(.TextMatrix(c, 6))
                RSStockUnit![Incoming] = toNumber(.TextMatrix(c, 7))
                RSStockUnit![Onhand] = toNumber(.TextMatrix(c, 8))

                RSStockUnit.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSStockUnit = Nothing
    
    CN.CommitTrans
  
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
    
    Exit Sub
err:
    If err.Number = -2147217887 Then
        Resume Next
    Else
        CN.RollbackTrans
        prompt_err err, Name, "cmdSave_Click"
        Screen.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdUsrHistory_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tDate2 As String
    Dim tUser1 As String
    Dim tUser2 As String
    
    tDate1 = Format$(RS.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    tDate2 = Format$(RS.Fields("DateModified"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("AddedByFK"), "CompleteName")
    tUser2 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & RS.Fields("LastUserFK"), "CompleteName")
    
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

'Procedure used to generate PK
Private Sub GeneratePK()
  PK = getIndex("Stock")
End Sub

Private Sub Form_Load()
    InitGrid
    InitNSD
    
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Stocks WHERE StockID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    rs1.CursorLocation = adUseClient
    rs1.Open "SELECT * FROM qry_Stock_Unit WHERE StockID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Stocks_Category order by category asc", "Category", dcCategory, "CategoryID", True
    bind_dc "SELECT * FROM Unit order by unit asc", "Unit", dcReoderUnit, "UnitID", True
        
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        dcCategory.Text = ""
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmProducts.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(1).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = RS![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmProductsAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtOrder.Text = .TextMatrix(.RowSel, 1)
        txtQty.Text = .TextMatrix(.RowSel, 2)
        nsdUnit.Text = .TextMatrix(.RowSel, 3)
        nsdUnit.Tag = .TextMatrix(.RowSel, 9) 'Add tag coz boundtext is empty
        txtSalesPrice.Text = .TextMatrix(.RowSel, 4)
        txtSupplierPrice.Text = .TextMatrix(.RowSel, 5)
        txtPending.Text = .TextMatrix(.RowSel, 6)
        txtIncoming.Text = .TextMatrix(.RowSel, 7)
        txtOnHand.Text = .TextMatrix(.RowSel, 8)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 9) = "" Then '10 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub lvPriceHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With lvPriceHistory
    'MsgBox .ColumnHeaders(2).Width & vbCr _
    & .ColumnHeaders(3).Width & vbCr _
    & .ColumnHeaders(4).Width
  
  End With
End Sub

Private Sub nsdUnit_Change()
    nsdUnit.Tag = nsdUnit.BoundText
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 9 Or Index = 10 Then KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
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
        .Cols = 10
        .ColSel = 9
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 800
        .ColWidth(2) = 800
        .ColWidth(3) = 800
        .ColWidth(4) = 900
        .ColWidth(5) = 1200
        .ColWidth(6) = 1200
        .ColWidth(7) = 1200
        .ColWidth(8) = 900
        .ColWidth(9) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Order"
        .TextMatrix(0, 2) = "Qty"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Sales Price"
        .TextMatrix(0, 5) = "Supplier Price"
        .TextMatrix(0, 6) = "Pending"
        .TextMatrix(0, 7) = "Incoming"
        .TextMatrix(0, 8) = "On Hand"
        .TextMatrix(0, 9) = "Unit ID"
        
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
'        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = flexAlignGeneral
'        .ColAlignment(4) = flexAlignGeneral
'        .ColAlignment(5) = vbRightJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
    End With
End Sub

Private Sub InitNSD()
    'For Vendor
    With nsdUnit
        .ClearColumn
        .AddColumn "Unit ID", 1794.89
        .AddColumn "Unit", 2264.88
        .Connection = CN.ConnectionString
        
        '.sqlFields = "VendorID, Company, Location"
        .sqlFields = "UnitID, Unit"
        .sqlTables = "Unit"
        .sqlSortOrder = "Unit ASC"
        
        .BoundField = "UnitID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Units Record"
    End With
    
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsUnit As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    rsUnit.CursorLocation = adUseClient
    rsUnit.Open "SELECT * FROM Stock_Unit WHERE StockID=" & PK, CN, adOpenStatic, adLockOptimistic
    If rsUnit.RecordCount > 0 Then
        rsUnit.MoveFirst
        While Not rsUnit.EOF
            CurrRow = getFlexPos(Grid, 9, rsUnit!UnitID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Stock_Unit", "StockUnitID", "", True, rsUnit!StockUnitID
                End If
            End With
            rsUnit.MoveNext
        Wend
    End If
End Sub

Private Sub txtIncoming_GotFocus()
    HLText txtIncoming
End Sub

Private Sub txtIncoming_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtIncoming_Validate(Cancel As Boolean)
    txtIncoming.Text = toNumber(txtIncoming.Text)
End Sub

Private Sub txtOnHand_GotFocus()
    HLText txtOnHand
End Sub

Private Sub txtOnHand_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtOnHand_Validate(Cancel As Boolean)
    txtOnHand.Text = toNumber(txtOnHand.Text)
End Sub

Private Sub txtOrder_GotFocus()
    HLText txtOrder
End Sub

Private Sub txtOrder_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtPending_GotFocus()
    HLText txtPending
End Sub

Private Sub txtPending_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtPending_Validate(Cancel As Boolean)
    txtPending.Text = toNumber(txtPending.Text)
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtSalesPrice_GotFocus()
    HLText txtSalesPrice
End Sub

Private Sub txtSalesPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtSalesPrice_Validate(Cancel As Boolean)
    txtSalesPrice.Text = toMoney(toNumber(txtSalesPrice.Text))
End Sub

Private Sub txtSupplierPrice_GotFocus()
    HLText txtSupplierPrice
End Sub

Private Sub txtSupplierPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtSupplierPrice_Validate(Cancel As Boolean)
    txtSupplierPrice.Text = toMoney(toNumber(txtSupplierPrice.Text))
End Sub
