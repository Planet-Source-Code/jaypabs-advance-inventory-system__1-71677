VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmQtyAdjustmentAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Qty Adjustment"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   8955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Qty Adjustment Tasks"
      Height          =   315
      Left            =   5280
      TabIndex        =   24
      Top             =   5070
      Width           =   1785
   End
   Begin VB.ComboBox cboReason 
      Height          =   315
      ItemData        =   "frmQtyAdjustmentAE.frx":0000
      Left            =   1170
      List            =   "frmQtyAdjustmentAE.frx":000A
      TabIndex        =   20
      Top             =   90
      Width           =   2115
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   150
      ScaleHeight     =   630
      ScaleWidth      =   8715
      TabIndex        =   5
      Top             =   1410
      Width           =   8715
      Begin VB.TextBox txtOldQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4830
         Locked          =   -1  'True
         TabIndex        =   22
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6540
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.TextBox txtNewQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4125
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   8
         Top             =   210
         Width           =   4050
         _ExtentX        =   7144
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
         Left            =   5550
         TabIndex        =   9
         Top             =   240
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
         Caption         =   "Old Qty"
         Height          =   240
         Index           =   0
         Left            =   4830
         TabIndex        =   23
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   5550
         TabIndex        =   12
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
         TabIndex        =   11
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "New Qty"
         Height          =   240
         Index           =   10
         Left            =   4125
         TabIndex        =   10
         Top             =   0
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8055
      TabIndex        =   4
      Top             =   5070
      Width           =   795
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   7140
      TabIndex        =   3
      Top             =   5070
      Width           =   855
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmQtyAdjustmentAE.frx":001F
      Left            =   1170
      List            =   "frmQtyAdjustmentAE.frx":0029
      TabIndex        =   2
      Text            =   "On Hold"
      Top             =   870
      Width           =   2325
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   90
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Tag             =   "Remarks"
      Top             =   5340
      Width           =   4980
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   180
      Picture         =   "frmQtyAdjustmentAE.frx":0040
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Remove"
      Top             =   2520
      Visible         =   0   'False
      Width           =   275
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2370
      Width           =   8715
      _ExtentX        =   15372
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
      Left            =   1170
      TabIndex        =   14
      Top             =   480
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20578307
      CurrentDate     =   38207
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1170
      TabIndex        =   15
      Top             =   480
      Width           =   2505
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Reason"
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   1005
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   90
      TabIndex        =   18
      Top             =   5070
      Width           =   990
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   90
      X2              =   8900
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   90
      X2              =   8900
      Y1              =   1290
      Y2              =   1290
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
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
      Left            =   180
      TabIndex        =   17
      Top             =   2070
      Width           =   4365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date Issued:"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   510
      Width           =   1035
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   120
      Top             =   2070
      Width           =   8715
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "History"
      End
   End
End
Attribute VB_Name = "frmQtyAdjustmentAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Old txtQty Value. Hold when editing qty

Private Sub btnAdd_Click()
On Error GoTo erR
    
    Dim RSStockUnit As New Recordset
    
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub
    
    If dcUnit.Text = "" Then
        MsgBox "Please select unit", vbInformation
        dcUnit.SetFocus
        Exit Sub
    End If
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 5, nsdStock.Tag)
    intStockID = nsdStock.Tag
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & intStockID & " AND UnitID = " & dcUnit.BoundText & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                .TextMatrix(1, 1) = nsdStock.Text
                .TextMatrix(1, 2) = txtNewQty.Text
                .TextMatrix(1, 3) = txtOldQty.Text
                .TextMatrix(1, 4) = dcUnit.Text
                .TextMatrix(1, 5) = intStockID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.Text
                .TextMatrix(.Rows - 1, 2) = txtNewQty.Text
                .TextMatrix(.Rows - 1, 3) = txtOldQty.Text
                .TextMatrix(.Rows - 1, 4) = dcUnit.Text
                .TextMatrix(.Rows - 1, 5) = intStockID
                
                .FillStyle = 1

                .Row = .Rows - 1
                .ColSel = 3
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already exist. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                                
                .TextMatrix(CurrRow, 1) = nsdStock.Text
                .TextMatrix(CurrRow, 2) = txtNewQty.Text
                .TextMatrix(CurrRow, 3) = txtOldQty.Text
                .TextMatrix(CurrRow, 4) = dcUnit.Text
                
                    'restore qty to Stock Unit's table
                RSStockUnit!Onhand = RSStockUnit!Onhand - intQtyOld

                RSStockUnit.Update
            Else
                Exit Sub
            End If
        End If
               
            'Add/deduct qty from Stock Unit's table
        RSStockUnit!Onhand = txtNewQty.Text

        RSStockUnit.Update
            
        'Highlight the current row's column
        .ColSel = 4
        'Display a remove button
        
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
    
    Exit Sub
    
erR:
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub btnRemove_Click()
    If MsgBox("This will restore the qty added previously from Products profile" & vbCrLf & vbCrLf & "Are you sure you want to continue?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    Dim RSStockUnit As New Recordset
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & nsdStock.Tag & " AND UnitID = " & dcUnit.BoundText & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic

        'restore qty to Stock Unit's table
    RSStockUnit!Onhand = RSStockUnit!Onhand + (toNumber(Grid.TextMatrix(Grid.RowSel, 3)) - toNumber(Grid.TextMatrix(Grid.RowSel, 2)))

    RSStockUnit.Update
    
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With
    
    'Save to stock card
'    Dim RSStockCard As New Recordset
'
'    With RSStockCard
'        .CursorLocation = adUseClient
'        .Open "SELECT * FROM Stock_Card WHERE StockID = " & toNumber(Grid.TextMatrix(Grid.RowSel, 10)) & " AND RefNo2 = '" & txtRefNo.Text & "'", CN, adOpenStatic, adLockOptimistic
'
'        !Pieces2 = !Pieces2 - toNumber(Grid.TextMatrix(Grid.RowSel, 3))
'
'        .Update
'    End With
'
    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub CmdTasks_Click()
    PopupMenu mnu_Tasks
End Sub

Private Sub mnu_History_Click()
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

Private Sub cmdCancel_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo erR

    'Verify the entries
    If cboReason.Text = "" Then
        MsgBox "Please select a reason.", vbExclamation
        cboReason.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to Adjust before you can save this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
              
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Qty_Adjustment_Detail WHERE QtyAdjustmentID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    DeleteItems
    
    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![QtyAdjustmentID] = PK
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        ElseIf State = adStateEditMode Then
            .Close
            .Open "SELECT * FROM Qty_Adjustment WHERE QtyAdjustmentID=" & PK, CN, adOpenStatic, adLockOptimistic
            
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        
        !Reason = cboReason.Text
        !Date = dtpDate.Value
        ![Status] = IIf(cboStatus.Text = "Adjusted", True, False)
        ![Notes] = txtNotes.Text

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

                RSDetails![QtyAdjustmentID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 5))
                RSDetails![NewQty] = toNumber(.TextMatrix(c, 2))
                RSDetails![OldQty] = toNumber(.TextMatrix(c, 3))
                RSDetails![UnitID] = getUnitID(.TextMatrix(c, 4))
                
                RSDetails.Update
                
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 5))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew
                
'                RSDetails![QtyAdjustmentID] = PK
'                RSDetails![StockID] = toNumber(.TextMatrix(c, 5))
                RSDetails![NewQty] = toNumber(.TextMatrix(c, 2))
                RSDetails![OldQty] = toNumber(.TextMatrix(c, 3))
                RSDetails![UnitID] = getUnitID(.TextMatrix(c, 4))
                
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

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
        Else
            Unload Me
        End If
    Else
        MsgBox "Changes in record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
erR:
    blnSave = False
'    CN.RollbackTrans
'    CN.BeginTrans
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub dcUnit_Click(Area As Integer)
    Call GetQty
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        cboReason.SetFocus
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
    If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM Qty_Adjustment WHERE QtyAdjustmentID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDate.Value = Date
        
        CN.BeginTrans

        GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_Qty_Adjustment WHERE QtyAdjustmentID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If State = adStateViewMode Then
            cmdCancel.Caption = "Close"
                   
            DisplayForViewing
        Else
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
    PK = getIndex("Qty_Adjustment")
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtNewQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmQtyAdjustment.RefreshRecords
    End If
    
    Set frmQtyAdjustmentAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If State = adStateViewMode Then Exit Sub

        nsdStock.Text = .TextMatrix(.RowSel, 1)
        nsdStock.Tag = .TextMatrix(.RowSel, 5) 'Add tag coz boundtext is empty
        intQtyOld = IIf(.TextMatrix(.RowSel, 2) = "", 0, .TextMatrix(.RowSel, 2))
        txtNewQty.Text = .TextMatrix(.RowSel, 2)
        
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 5), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        dcUnit.Text = .TextMatrix(.RowSel, 4)
        'disable unit to prevent user from changing it. changing of unit will result to imbalance of inventory
        If State = adStateEditMode Then dcUnit.Enabled = False
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 5) = "" Then '5 = StockID
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
    
    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText, "Unit", dcUnit, "UnitID", True
    
    Call GetQty
End Sub

Private Sub GetQty()
    Dim RSStockUnit As New Recordset
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & nsdStock.BoundText & " AND UnitID = " & dcUnit.BoundText & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    'Retrieve qty from Stock Unit's table
    txtOldQty.Text = RSStockUnit!Onhand
    txtNewQty.Text = RSStockUnit!Onhand
End Sub

Private Sub txtNewQty_Validate(Cancel As Boolean)
    txtNewQty.Text = toNumber(txtNewQty.Text)
End Sub

Private Sub txtNewQty_Change()
    If toNumber(txtNewQty.Text) < 1 Then
        btnAdd.Enabled = False
        Exit Sub
    Else
        btnAdd.Enabled = True
    End If
End Sub

Private Sub txtNewQty_GotFocus()
    HLText txtNewQty
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
    On Error GoTo erR
    
    cboReason.Text = rs![Reason]
    dtpDate.Value = rs![Date]
    cboStatus.Text = rs!Status_Alias
    txtNotes.Text = rs![Notes]
    
    cIRowCount = 0

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Qty_Adjustment_Detail WHERE QtyAdjustmentID=" & PK & " ORDER BY QtyAdjDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Stock]
                    .TextMatrix(1, 2) = RSDetails![NewQty]
                    .TextMatrix(1, 3) = RSDetails![OldQty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![NewQty]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![OldQty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 4
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionByRow
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
   
    Exit Sub
erR:
    'Error if encounter a null value
    If erR.Number = 94 Then
        Resume Next
    Else
        MsgBox erR.Description
    End If
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo erR
    
    cboReason.Text = rs![Reason]
    dtpDate.Value = rs![Date]
    cboStatus.Text = rs!Status_Alias
    txtNotes.Text = rs![Notes]

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Qty_Adjustment_Detail WHERE QtyAdjustmentID=" & PK & " ORDER BY QtyAdjDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Stock]
                    .TextMatrix(1, 2) = RSDetails![NewQty]
                    .TextMatrix(1, 3) = RSDetails![OldQty]
                    .TextMatrix(1, 4) = RSDetails![Unit]
                    .TextMatrix(1, 5) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![NewQty]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![OldQty]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 5) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 4
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
    btnAdd.Visible = False

    'Resize and reposition the controls
   
    Shape3.Top = 1300
    Label11.Top = 1300
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 1600
    Grid.Height = 3280

    Exit Sub
erR:
    'Error if encounter a null value
    If erR.Number = 94 Then
        Resume Next
    Else
        MsgBox erR.Description
    End If
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
        .Cols = 6
        .ColSel = 5
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 3525
        .ColWidth(2) = 1545
        .ColWidth(3) = 1545
        .ColWidth(4) = 1545
        .ColWidth(5) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "New Qty"
        .TextMatrix(0, 3) = "Old Qty"
        .TextMatrix(0, 4) = "Unit"
        .TextMatrix(0, 5) = "Stock ID"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbLeftJustify
    End With
End Sub

Private Sub InitNSD()
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

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSQtyAdj As New Recordset

    If State = adStateAddMode Then Exit Sub

    RSQtyAdj.CursorLocation = adUseClient
    RSQtyAdj.Open "SELECT * FROM Qty_Adjustment_Detail WHERE QtyAdjustmentID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSQtyAdj.RecordCount > 0 Then
        RSQtyAdj.MoveFirst
        While Not RSQtyAdj.EOF
            CurrRow = getFlexPos(Grid, 5, RSQtyAdj!StockID)

            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Qty_Adjustment_Detail", "QtyAdjDetailID", "", True, RSQtyAdj!QtyAdjDetailID
                End If
            End With
            RSQtyAdj.MoveNext
        Wend
    End If
End Sub
