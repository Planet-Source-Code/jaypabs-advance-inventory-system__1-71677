VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSalesReceiptsBatch 
   Caption         =   "Sales Receipts by Batch"
   ClientHeight    =   8325
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   10365
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture3 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10365
      TabIndex        =   13
      Top             =   7920
      Width           =   10365
      Begin VB.PictureBox Picture4 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5700
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   14
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnFirst2 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev2 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   17
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast2 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnNext2 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   15
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo2 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.Label lblCurrentRecord2 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   60
         Width           =   1365
      End
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   1
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10365
      TabIndex        =   9
      Top             =   8295
      Width           =   10365
   End
   Begin VB.PictureBox picLine 
      Align           =   2  'Align Bottom
      BackColor       =   &H80000014&
      BorderStyle     =   0  'None
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   10365
      TabIndex        =   8
      Top             =   8310
      Width           =   10365
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   0
      ScaleHeight     =   375
      ScaleWidth      =   10365
      TabIndex        =   0
      Top             =   3690
      Width           =   10365
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Delete"
         Height          =   315
         Left            =   4680
         TabIndex        =   24
         Top             =   30
         Width           =   1185
      End
      Begin VB.CommandButton cmdReturn 
         Caption         =   "&Return"
         Height          =   315
         Left            =   5940
         TabIndex        =   23
         Top             =   30
         Width           =   1185
      End
      Begin VB.CommandButton CmdAddReceipt 
         Caption         =   "&Add Receipt"
         Height          =   315
         Left            =   3390
         TabIndex        =   22
         Top             =   30
         Width           =   1185
      End
      Begin VB.CommandButton cmdPrint 
         Caption         =   "&Print"
         Height          =   315
         Left            =   1980
         TabIndex        =   21
         Top             =   30
         Width           =   1335
      End
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   5700
         ScaleHeight     =   345
         ScaleWidth      =   4155
         TabIndex        =   1
         Top             =   0
         Width           =   4150
         Begin VB.CommandButton btnNext1 
            Height          =   315
            Left            =   3390
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Next 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnLast1 
            Height          =   315
            Left            =   3705
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Last 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnPrev1 
            Height          =   315
            Left            =   3075
            Style           =   1  'Graphical
            TabIndex        =   3
            ToolTipText     =   "Previous 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.CommandButton btnFirst1 
            Height          =   315
            Left            =   2760
            Style           =   1  'Graphical
            TabIndex        =   2
            ToolTipText     =   "First 250"
            Top             =   10
            Width           =   315
         End
         Begin VB.Label lblPageInfo1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "0 - 0 of 0"
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   60
            Width           =   2535
         End
      End
      Begin VB.Label lblCurrentRecord1 
         AutoSize        =   -1  'True
         Caption         =   "Selected Record: 0"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   60
         Width           =   1365
      End
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   2835
      Left            =   0
      TabIndex        =   10
      Top             =   390
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5001
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Route"
         Object.Width           =   3246
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "TruckNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Booking"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Collection"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Delivery Date"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lvList2 
      Height          =   3225
      Left            =   0
      TabIndex        =   12
      Top             =   4200
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   5689
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   8
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Company"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "City"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "RefNo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Agent"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Date Of Delivery"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Status"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Deducted"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Printed"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Receipts"
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
      Left            =   75
      TabIndex        =   11
      Top             =   120
      Width           =   4815
   End
   Begin VB.Shape shpBar 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   0
      Top             =   120
      Width           =   6915
   End
End
Attribute VB_Name = "frmSalesReceiptsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim CURR_COL            As Integer
Dim RSReceiptsBatch     As New Recordset
Dim RSReceipts          As New Recordset
Dim RecordPageBatch     As New clsPaging
Dim SQLParserBatch      As New clsSQLSelectParser
Dim RecordPage          As New clsPaging
Dim SQLParser           As New clsSQLSelectParser
Dim intRoute            As Integer

'Procedure used to filter records
Public Sub FilterRecord(ByVal srcCondition As String)
    SQLParserBatch.RestoreStatement
    SQLParserBatch.wCondition = srcCondition
    
    ReloadRecords1 SQLParserBatch.SQLStatement
End Sub

Public Sub CommandPass(ByVal srcPerformWhat As String)
    On Error GoTo err
    Select Case srcPerformWhat
        Case "New"
            frmSalesReceiptsBatchAE.State = adStateAddMode
            frmSalesReceiptsBatchAE.show vbModal
        Case "Edit"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("Receipts_Batch", "ReceiptBatchID", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user. Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords1
                    Exit Sub
                Else
                    With frmSalesReceiptsBatchAE
                        .State = adStateEditMode
                        .PK = CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        .show vbModal
                        RefreshRecords1
                    End With
                End If
            End If
        Case "Search"
            With frmSearch
                Set .srcForm = Me
                Set .srcColumnHeaders = lvList.ColumnHeaders
                .show vbModal
            End With
        Case "Delete"
            If lvList.ListItems.Count > 0 Then
                If isRecordExist("Receipts_Batch", "ReceiptBatchID", CLng(LeftSplitUF(lvList.SelectedItem.Tag))) = False Then
                    MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
                    RefreshRecords1
                    Exit Sub
                Else
                    Dim ANS As Integer
                    ANS = MsgBox("Are you sure you want to void the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation. This will permanently remove the record.", vbCritical + vbYesNo, "Confirm Record")
                    Me.MousePointer = vbHourglass
                    If ANS = vbYes Then
                        'Remove
                        DelRecwSQL "Receipts_Batch", "ReceiptBatchID", "", True, CLng(LeftSplitUF(lvList.SelectedItem.Tag))
                        'Refresh the records
                        RefreshRecords1
                        MsgBox "Record has been successfully removed.", vbInformation, "Confirm"
                    End If
                    ANS = 0
                    Me.MousePointer = vbDefault
                End If
            Else
                MsgBox "No record to void.", vbExclamation
            End If
        Case "Refresh"
            RefreshRecords1
        Case "Print"
        Case "Close"
            Unload Me
    End Select
    Exit Sub
    'Trap the error
err:
    If err.Number = -2147467259 Then
        MsgBox "You cannot delete this record because it was used by other records! If you want to delete this record" & vbCrLf & _
               "you will first have to delete or change the records that currenly used this record as shown bellow." & vbCrLf & vbCrLf & _
               err.Description, , "Delete Operation Failed!"
        Me.MousePointer = vbDefault
    Else
        MsgBox err.Description, vbInformation
        Me.MousePointer = vbDefault
    End If
End Sub

Public Sub RefreshRecords1()
    SQLParserBatch.RestoreStatement
    ReloadRecords1 SQLParserBatch.SQLStatement
End Sub

Public Sub RefreshRecords2()
    SQLParser.RestoreStatement
    ReloadRecords2 SQLParser.SQLStatement
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords1(ByVal srcSQL As String)
    '-In this case I used SQL because it is faster than Filter function of VB
    '-when hundling millions of records.
    On Error GoTo err
    With RSReceiptsBatch
        If .State = adStateOpen Then .Close
        .Open srcSQL
    End With
    RecordPageBatch.Refresh
    FillList1 1
    Exit Sub
err:
        If err.Number = -2147217913 Then
            srcSQL = Replace(srcSQL, "'", "", , , vbTextCompare)
            Resume
        ElseIf err.Number = -2147217900 Then
            MsgBox "Invalid search operation.", vbExclamation
            SQLParserBatch.RestoreStatement
            srcSQL = SQLParserBatch.SQLStatement
            Resume
        Else
            prompt_err err, Name, "ReloadRecords1"
        End If
End Sub

'Procedure for reloadingrecords
Public Sub ReloadRecords2(ByVal srcSQL As String)
    '-In this case I used SQL because it is faster than Filter function of VB
    '-when hundling millions of records.
    On Error GoTo err
    With RSReceipts
        If .State = adStateOpen Then .Close
        .Open srcSQL
    End With
    RecordPage.Refresh
    FillList2 1
    Exit Sub
err:
        If err.Number = -2147217913 Then
            srcSQL = Replace(srcSQL, "'", "", , , vbTextCompare)
            Resume
        ElseIf err.Number = -2147217900 Then
            MsgBox "Invalid search operation.", vbExclamation
            SQLParser.RestoreStatement
            srcSQL = SQLParser.SQLStatement
            Resume
        Else
            prompt_err err, Name, "ReloadRecords2"
        End If
End Sub

Private Sub btnClose_Click()
    Unload Me
End Sub

Private Sub btnFirst1_Click()
    If RecordPageBatch.PAGE_CURRENT <> 1 Then FillList1 1
End Sub

Private Sub btnLast1_Click()
    If RecordPageBatch.PAGE_CURRENT <> RecordPageBatch.PAGE_TOTAL Then FillList1 RecordPageBatch.PAGE_TOTAL
End Sub

Private Sub btnNext1_Click()
    If RecordPageBatch.PAGE_CURRENT <> RecordPageBatch.PAGE_TOTAL Then FillList1 RecordPageBatch.PAGE_NEXT
End Sub

Private Sub btnPrev1_Click()
    If RecordPageBatch.PAGE_CURRENT <> 1 Then FillList1 RecordPageBatch.PAGE_PREVIOUS
End Sub

Private Sub CmdAddReceipt_Click()
    With frmSalesReceiptsAE
        .State = adStateAddMode
        .ReceiptBatchPK = LeftSplitUF(lvList.SelectedItem.Tag)
        .dtpDeliveryDate = lvList.SelectedItem.SubItems(4)
        .dcAgent = lvList.SelectedItem.SubItems(2)
        .show vbModal
        
        RefreshRecords1
    End With
End Sub

Private Sub cmdDelete_Click()
    If lvList.ListItems.Count > 0 Then
        If isRecordExist("Receipts", "ReceiptID", CLng(LeftSplitUF(lvList2.SelectedItem.Tag))) = False Then
            MsgBox "This record has been removed by other user.Click 'OK' button to refresh the records.", vbExclamation
            RefreshRecords1
            Exit Sub
        Else
            Dim ANS As Integer
            ANS = MsgBox("Are you sure you want to void the selected record?" & vbCrLf & vbCrLf & "WARNING: You cannot undo this operation. This will permanently remove the record.", vbCritical + vbYesNo, "Confirm Record")
            Me.MousePointer = vbHourglass
            If ANS = vbYes Then
                'Remove
                DelRecwSQL "Receipts", "ReceiptID", "", True, CLng(LeftSplitUF(lvList2.SelectedItem.Tag))
                'Refresh the records
                RefreshRecords2
                MsgBox "Record has been successfully removed.", vbInformation, "Confirm"
            End If
            ANS = 0
            Me.MousePointer = vbDefault
        End If
    Else
        MsgBox "No record to void.", vbExclamation
    End If
End Sub

Private Sub cmdPrint_Click()
    PopupMenu MAIN.mnu_ReceiptsBatch
End Sub

Private Sub cmdReturn_Click()
    Dim RSSalesReturn As New Recordset
    Dim ReceiptPK As Integer
    
    ReceiptPK = CLng(LeftSplitUF(lvList2.SelectedItem.Tag))
    
    RSSalesReturn.CursorLocation = adUseClient
    RSSalesReturn.Open "SELECT SalesReturnID FROM Sales_Return WHERE ReceiptID=" & ReceiptPK, CN, adOpenStatic, adLockOptimistic
    
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
            .ReceiptPK = ReceiptPK
        End If
        
        .show vbModal
    End With
End Sub

Private Sub Form_Activate()
    HighlightInWin Me.Name: MAIN.ShowTBButton "tttttft"
    Active
End Sub

Private Sub Form_Deactivate()
    MAIN.HideTBButton "", True
    Deactive
End Sub

Private Sub Active()
    With MAIN
        .tbMenu.Buttons(4).Caption = "View"
        .tbMenu.Buttons(6).Caption = "Void"
        .tbMenu.Buttons(4).Image = 13
        .tbMenu.Buttons(6).Image = 14

        .mnuRAES.Caption = "View Selected"
        .mnuRADS.Caption = "Void Selected"
    End With
End Sub

Private Sub Deactive()
    With MAIN
        .tbMenu.Buttons(4).Caption = "Edit"
        .tbMenu.Buttons(6).Caption = "Delete"
        .tbMenu.Buttons(4).Image = 2
        .tbMenu.Buttons(6).Image = 4

        .mnuRAES.Caption = "Edit Selected"
        .mnuRADS.Caption = "Delete Selected"
    End With
End Sub

Private Sub Form_Load()
    MAIN.AddToWin Me.Caption, Name

    'Set the graphics for the controls
    With MAIN
        'For listview
        Set lvList.SmallIcons = .i16x16
        Set lvList.Icons = .i16x16
    
        Set lvList2.SmallIcons = .i16x16
        Set lvList2.Icons = .i16x16
    
        btnFirst1.Picture = .i16x16.ListImages(3).Picture
        btnPrev1.Picture = .i16x16.ListImages(4).Picture
        btnNext1.Picture = .i16x16.ListImages(5).Picture
        btnLast1.Picture = .i16x16.ListImages(6).Picture
        
        btnFirst2.Picture = .i16x16.ListImages(3).Picture
        btnPrev2.Picture = .i16x16.ListImages(4).Picture
        btnNext2.Picture = .i16x16.ListImages(5).Picture
        btnLast2.Picture = .i16x16.ListImages(6).Picture
        
        btnFirst1.DisabledPicture = .i16x16g.ListImages(3).Picture
        btnPrev1.DisabledPicture = .i16x16g.ListImages(4).Picture
        btnNext1.DisabledPicture = .i16x16g.ListImages(5).Picture
        btnLast1.DisabledPicture = .i16x16g.ListImages(6).Picture
    
        btnFirst2.DisabledPicture = .i16x16g.ListImages(3).Picture
        btnPrev2.DisabledPicture = .i16x16g.ListImages(4).Picture
        btnNext2.DisabledPicture = .i16x16g.ListImages(5).Picture
        btnLast2.DisabledPicture = .i16x16g.ListImages(6).Picture
    End With
    
    With SQLParser
        .Fields = "Company, City, RefNo, AgentName, DateIssued, Status_Alias, Deducted, Printed_Alias, ReceiptID"
        .Tables = "qry_Receipts"
        .SortOrder = "DateIssued DESC"
        
        .SaveStatement
    End With
    
    RSReceipts.CursorLocation = adUseClient
    RSReceipts.Open SQLParser.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
'    With RecordPage
'        .Start RSReceipts, 75
'        FillList2 1
'    End With
    
    With SQLParserBatch
        .Fields = "Desc, TruckNo, Booking, Collection, DeliveryDate, Status_Alias, ReceiptBatchID"
        .Tables = "qry_Receipts_Batch"
        .SortOrder = "DeliveryDate DESC"
        
        .SaveStatement
    End With
    
    RSReceiptsBatch.CursorLocation = adUseClient
    RSReceiptsBatch.Open SQLParserBatch.SQLStatement, CN, adOpenStatic, adLockReadOnly
    
    With RecordPageBatch
        .Start RSReceiptsBatch, 75
        FillList1 1
    End With
End Sub

Private Sub FillList1(ByVal whichPage As Long)
    RecordPageBatch.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList, RSReceiptsBatch, RecordPageBatch.PageStart, RecordPageBatch.PageEnd, 16, 2, False, True, , , , "ReceiptBatchID")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation1
    'Display the page information
    lblPageInfo1.Caption = "Record " & RecordPageBatch.PageInfo
    'Display the selected record
    lvList_Click
End Sub

Private Sub FillList2(ByVal whichPage As Long)
    RecordPage.CurrentPosition = whichPage
    Screen.MousePointer = vbHourglass
    Me.Enabled = False
    Call pageFillListView(lvList2, RSReceipts, RecordPage.PageStart, RecordPage.PageEnd, 16, 2, False, True, , , , "ReceiptID")
    Me.Enabled = True
    Screen.MousePointer = vbDefault
    SetNavigation2
    'Display the page information
    lblPageInfo2.Caption = "Record " & RecordPage.PageInfo
    'Display the selected record
    lvList2_Click
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If WindowState <> vbMinimized Then
        If Me.Width < 9195 Then Me.Width = 9195
        If Me.Height < 4500 Then Me.Height = 4500
        
        shpBar.Width = ScaleWidth
        
        lvList.Width = Me.ScaleWidth
        lvList.Height = Me.Height / 2 - 1500
        
        Picture1.Top = lvList.Height - lvList.Top + 1000
        
        lvList2.Top = lvList.Height - lvList.Top + 1500
        lvList2.Width = Me.ScaleWidth
        lvList2.Height = Me.Height / 2 - 600
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    MAIN.RemToWin Me.Caption
    MAIN.HideTBButton "", True
    
    Set frmSalesReceiptsBatch = Nothing
End Sub

Private Sub SetNavigation1()
    With RecordPageBatch
        If .PAGE_TOTAL = 1 Then
            btnFirst1.Enabled = False
            btnPrev1.Enabled = False
            btnNext1.Enabled = False
            btnLast1.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            btnFirst1.Enabled = False
            btnPrev1.Enabled = False
            btnNext1.Enabled = True
            btnLast1.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            btnFirst1.Enabled = True
            btnPrev1.Enabled = True
            btnNext1.Enabled = False
            btnLast1.Enabled = False
        Else
            btnFirst1.Enabled = True
            btnPrev1.Enabled = True
            btnNext1.Enabled = True
            btnLast1.Enabled = True
        End If
    End With
End Sub

Private Sub SetNavigation2()
    With RecordPage
        If .PAGE_TOTAL = 1 Then
            btnFirst2.Enabled = False
            btnPrev2.Enabled = False
            btnNext2.Enabled = False
            btnLast2.Enabled = False
        ElseIf .PAGE_CURRENT = 1 Then
            btnFirst2.Enabled = False
            btnPrev2.Enabled = False
            btnNext2.Enabled = True
            btnLast2.Enabled = True
        ElseIf .PAGE_CURRENT = .PAGE_TOTAL And .PAGE_CURRENT > 1 Then
            btnFirst2.Enabled = True
            btnPrev2.Enabled = True
            btnNext2.Enabled = False
            btnLast2.Enabled = False
        Else
            btnFirst2.Enabled = True
            btnPrev2.Enabled = True
            btnNext2.Enabled = True
            btnLast2.Enabled = True
        End If
    End With
End Sub

Private Sub lvList_Click()
    On Error GoTo err
    lblCurrentRecord1.Caption = "Selected Record: " & RightSplitUF(lvList.SelectedItem.Tag)
        
    SQLParser.RestoreStatement
    'SQLParser.wCondition = "DateofDelivery = #" & Format(lvList.ListItems(lvList.SelectedItem.Index), "Short Date") & "# AND Route = " & lvList.ListItems(1).SubItems(1) & ""
    SQLParser.wCondition = "ReceiptBatchID = " & LeftSplitUF(lvList.SelectedItem.Tag)
    
    With RecordPage
        .Start RSReceipts, 75
        FillList2 1
    End With
    
    frmSalesReceiptsAE.strRouteDesc = lvList.SelectedItem.Text
    
    ReloadRecords2 SQLParser.SQLStatement
        
    Exit Sub
err:
    lblCurrentRecord1.Caption = "Selected Record: NONE"
End Sub

Private Sub lvList_DblClick()
    CommandPass "Edit"
End Sub

Private Sub lvList_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu MAIN.mnuRecA
End Sub

Private Sub lvList_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sort the listview
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lvList.SortOrder = 0
    Else
        lvList.SortOrder = Abs(lvList.SortOrder - 1)
    End If
    lvList.SortKey = ColumnHeader.Index - 1
    
    lvList.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub lvList_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Or KeyCode = 40 Or KeyCode = 33 Or KeyCode = 34 Then lvList_Click
End Sub

Private Sub lvList2_Click()
    On Error GoTo err
    lblCurrentRecord2.Caption = "Selected Record: " & RightSplitUF(lvList2.SelectedItem.Tag)
        
    Exit Sub
err:
        lblCurrentRecord2.Caption = "Selected Record: NONE"
End Sub

Private Sub lvList2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    'Sort the listview
    If ColumnHeader.Index - 1 <> CURR_COL Then
        lvList2.SortOrder = 0
    Else
        lvList2.SortOrder = Abs(lvList2.SortOrder - 1)
    End If
    lvList2.SortKey = ColumnHeader.Index - 1
    
    lvList2.Sorted = True
    CURR_COL = ColumnHeader.Index - 1
End Sub

Private Sub lvList2_DblClick()
    If lvList2.ListItems.Count > 0 Then
        If isRecordExist("Receipts", "ReceiptID", CLng(LeftSplitUF(lvList2.SelectedItem.Tag))) = False Then
            MsgBox "This record has been removed by other user. Click 'OK' button to refresh the records.", vbExclamation
            RefreshRecords2
            Exit Sub
        Else
            With frmSalesReceiptsAE
                Dim blnStatus As Boolean
                
                blnStatus = getValueAt("SELECT ReceiptID,Status FROM Receipts WHERE ReceiptID=" & CLng(LeftSplitUF(lvList2.SelectedItem.Tag)), "Status")
                
                If blnStatus Then 'true
                    .State = adStateViewMode
                Else
                    .State = adStateEditMode
                End If
            
                .PK = CLng(LeftSplitUF(lvList2.SelectedItem.Tag))
                .ReceiptBatchPK = LeftSplitUF(lvList.SelectedItem.Tag)
                .show vbModal
                RefreshRecords1
            
'                .State = adStateEditMode
'                .PK = CLng(LeftSplitUF(lvList2.SelectedItem.Tag))
'                .show vbModal
'                RefreshRecords1
            End With
        End If
    End If
End Sub

Private Sub Picture3_Resize()
    Picture2.Left = Picture3.ScaleWidth - Picture4.ScaleWidth
    Picture4.Left = Picture3.ScaleWidth - Picture4.ScaleWidth
    
    Picture1.Width = Picture3.Width
End Sub
