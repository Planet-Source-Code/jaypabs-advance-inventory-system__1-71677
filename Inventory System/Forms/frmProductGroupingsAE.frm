VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmProductGroupingsAE 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7005
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7020
   LinkTopic       =   "Form1"
   ScaleHeight     =   7005
   ScaleWidth      =   7020
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   1065
      Index           =   3
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   5190
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Product Groupings Detail"
      Height          =   3315
      Left            =   210
      TabIndex        =   8
      Top             =   1470
      Width           =   6525
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   4080
         TabIndex        =   3
         Top             =   510
         Width           =   765
      End
      Begin VB.CommandButton btnRemove 
         Height          =   275
         Left            =   180
         Picture         =   "frmProductGroupingsAE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Remove"
         Top             =   1110
         Visible         =   0   'False
         Width           =   275
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   2190
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   3863
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
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   150
         TabIndex        =   2
         Top             =   495
         Width           =   3840
         _ExtentX        =   6773
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
         Left            =   150
         TabIndex        =   16
         Top             =   270
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   3930
      TabIndex        =   6
      Top             =   6450
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   5400
      TabIndex        =   7
      Top             =   6450
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   735
      Width           =   4065
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   2
      Left            =   1620
      MaxLength       =   200
      TabIndex        =   1
      Top             =   1080
      Width           =   1635
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   6480
      Width           =   1680
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   210
      TabIndex        =   11
      Top             =   6360
      Width           =   6465
      _ExtentX        =   18283
      _ExtentY        =   53
   End
   Begin VB.Shape Shape3 
      Height          =   6885
      Left            =   60
      Top             =   60
      Width           =   6885
   End
   Begin VB.Shape Shape1 
      Height          =   6255
      Left            =   150
      Top             =   630
      Width           =   6705
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   15
      Top             =   735
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Minimum Qty"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   14
      Top             =   1080
      Width           =   1245
   End
   Begin VB.Label Label7 
      Caption         =   "Remarks"
      Height          =   225
      Left            =   210
      TabIndex        =   13
      Top             =   4920
      Width           =   1095
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Product Groupings"
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
      Left            =   330
      TabIndex        =   12
      Top             =   180
      Width           =   4905
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   150
      Width           =   6705
   End
End
Attribute VB_Name = "frmProductGroupingsAE"
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
Dim blnRemarks              As Boolean
Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset
Dim rsProductGroupings As New Recordset

Private Sub btnAdd_Click()
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub

    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 2, nsdStock.Tag)

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 2) = "" Then
                .TextMatrix(1, 1) = nsdStock.Text
                .TextMatrix(1, 2) = nsdStock.Tag
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.Text
                .TextMatrix(.Rows - 1, 2) = nsdStock.Tag

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = nsdStock.Text
                .TextMatrix(CurrRow, 2) = nsdStock.Tag
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .ColSel = 1
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub DisplayForEditing()
    On Error GoTo erR
    Dim rsClients As New Recordset
    
    rsClients.CursorLocation = adUseClient
    rsClients.Open "SELECT * FROM Product_Groupings WHERE ProductGroupingID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    With rsClients
        txtEntry(1).Text = .Fields("Description")
        txtEntry(2).Text = .Fields("Qty")
        txtEntry(3).Text = .Fields("Notes")
    End With
    
    'Display the details
    Dim rsProductGroupings As New Recordset

    cIRowCount = 0
    
    rsProductGroupings.CursorLocation = adUseClient
    rsProductGroupings.Open "SELECT * FROM qry_Product_Groupings_Detail WHERE ProductGroupingID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsProductGroupings.RecordCount > 0 Then
        rsProductGroupings.MoveFirst
        While Not rsProductGroupings.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 2) = "" Then
                    .TextMatrix(1, 1) = rsProductGroupings![Stock]
                    .TextMatrix(1, 2) = rsProductGroupings![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsProductGroupings![Stock]
                    .TextMatrix(.Rows - 1, 2) = rsProductGroupings![StockID]
                End If
            End With
            rsProductGroupings.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 2
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsProductGroupings.Close
    'Clear variables
    Set rsProductGroupings = Nothing
        
    'txtEntry(1).SetFocus
    Exit Sub
erR:
    prompt_err erR, Name, "DisplayForEditing"
    Screen.MousePointer = vbDefault
End Sub

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

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
  clearText Me
  
  txtEntry(1).SetFocus
End Sub

Private Sub cmdSave_Click()
    On Error GoTo erR

    If Trim(txtEntry(1).Text) = "" Then Exit Sub
    
    CN.BeginTrans

    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("ProductGroupingID") = PK
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
      .Fields("Description") = txtEntry(1).Text
      .Fields("Qty") = txtEntry(2).Text
      .Fields("Notes") = txtEntry(3).Text
            
      .Update
    End With

    Dim rsProductGroupings As New Recordset

    rsProductGroupings.CursorLocation = adUseClient
    rsProductGroupings.Open "SELECT * FROM Product_Groupings_Detail WHERE ProductGroupingID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                rsProductGroupings.AddNew

                rsProductGroupings![ProductGroupingID] = PK
                rsProductGroupings![StockID] = toNumber(.TextMatrix(c, 2))

                rsProductGroupings.Update
            ElseIf State = adStateEditMode Then
                rsProductGroupings.Filter = "StockID = " & toNumber(.TextMatrix(c, 2))
            
                If rsProductGroupings.RecordCount = 0 Then GoTo AddNew

                rsProductGroupings![ProductGroupingID] = PK
                rsProductGroupings![StockID] = toNumber(.TextMatrix(c, 2))

                rsProductGroupings.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set rsProductGroupings = Nothing
    
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

erR:
    CN.RollbackTrans
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
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
    InitNSD
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Product_Groupings WHERE ProductGroupingID = " & PK, CN, adOpenStatic, adLockOptimistic
        
'    rsProductGroupings.CursorLocation = adUseClient
'    rsProductGroupings.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
   
    'Check the form state
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
    PK = getIndex("Product_Groupings")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmProductGroupings.RefreshRecords
        End If
    End If
    
    Set frmProductGroupingsAE = Nothing
End Sub

Private Sub ResetEntry()
    nsdStock.Text = ""
End Sub

Private Sub Grid_Click()
    With Grid
        nsdStock.Text = .TextMatrix(.RowSel, 1)
        nsdStock.Tag = .TextMatrix(.RowSel, 2) 'Add tag coz boundtext is empty
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 2) = "" Then '2 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub nsdStock_Change()
    nsdStock.Tag = nsdStock.BoundText
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 3 Then
        blnRemarks = True
        Exit Sub
    Else
        blnRemarks = False
    End If
    
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 2 Then KeyAscii = isNumber(KeyAscii)
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
        .Cols = 3
        .ColSel = 2
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 3400

        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Product"
        .TextMatrix(0, 2) = "Stock ID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
    End With
End Sub

Private Sub InitNSD()
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Stock", 4085.26
        .AddColumn "StockID", 0
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Stock, StockID"
        .sqlTables = "Stocks"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 1
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Products"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim rsPGDetails As New Recordset 'Product Groupings Detail
    
    If State = adStateAddMode Then Exit Sub
    
    rsPGDetails.CursorLocation = adUseClient
    rsPGDetails.Open "SELECT * FROM Product_Groupings_Detail WHERE ProductGroupingID=" & PK, CN, adOpenStatic, adLockOptimistic
    If rsPGDetails.RecordCount > 0 Then
        rsPGDetails.MoveFirst
        While Not rsPGDetails.EOF
            CurrRow = getFlexPos(Grid, 2, rsPGDetails!StockID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Product_Groupings_Detail", "ProductGroupingDetailID", "", True, rsPGDetails!ProductGroupingDetailID
                End If
            End With
            rsPGDetails.MoveNext
        Wend
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    If Index = 2 Then
        txtEntry(2).Text = toNumber(txtEntry(2).Text)
    End If
End Sub
