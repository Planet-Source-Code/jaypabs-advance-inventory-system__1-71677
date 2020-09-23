VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSalesExpenses 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Expenses"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6870
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "&Update"
      Height          =   315
      Left            =   3780
      TabIndex        =   9
      Top             =   600
      Width           =   1035
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   2220
      TabIndex        =   4
      Top             =   600
      Width           =   1425
   End
   Begin VB.TextBox txtDescription 
      Height          =   285
      Left            =   300
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   1815
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   450
      TabIndex        =   2
      Top             =   3630
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4335
      TabIndex        =   1
      Top             =   3615
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   3600
      Width           =   1335
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   270
      TabIndex        =   5
      Top             =   3480
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2190
      Left            =   270
      TabIndex        =   8
      Top             =   1020
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
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   285
      Left            =   2520
      TabIndex        =   7
      Top             =   210
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Description"
      Height          =   285
      Left            =   330
      TabIndex        =   6
      Top             =   240
      Width           =   1095
   End
End
Attribute VB_Name = "frmSalesExpenses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Dim cIRowCount              As Integer

Private Sub DisplayForEditing()
    On Error GoTo err
    
    'Display the details
    Dim rsSalesExpenses As New Recordset

    cIRowCount = 0
    
    rsSalesExpenses.CursorLocation = adUseClient
    rsSalesExpenses.Open "SELECT * FROM qry_Sales_Expenses WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsSalesExpenses.RecordCount > 0 Then
        rsSalesExpenses.MoveFirst
        While Not rsSalesExpenses.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 3) = "" Then
                    .TextMatrix(1, 1) = rsSalesExpenses![Description]
                    .TextMatrix(1, 2) = rsSalesExpenses![Amount]
                    .TextMatrix(1, 3) = rsSalesExpenses![DescriptionID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsSalesExpenses![Description]
                    .TextMatrix(.Rows - 1, 2) = rsSalesExpenses![Amount]
                    .TextMatrix(.Rows - 1, 3) = rsSalesExpenses![DescriptionID]
                End If
            End With
            rsSalesExpenses.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 2
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsSalesExpenses.Close
    'Clear variables
    Set rsSalesExpenses = Nothing
        
    'txtEntry(1).SetFocus
    Exit Sub
err:
    If err.Number = 94 Then
        Resume Next
    Else
        MsgBox err.Number & " " & err.Description, vbInformation
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
        .Cols = 4
        .ColSel = 3
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1400
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Amount"
        .TextMatrix(0, 3) = "DescriptionID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err
    
    Dim rsSalesExpenses As New Recordset
    
    rsSalesExpenses.CursorLocation = adUseClient
    rsSalesExpenses.Open "SELECT * FROM qry_Sales_Expenses WHERE ReceiptBatchID=" & PK, CN, adOpenStatic, adLockOptimistic

    CN.BeginTrans
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c

            rsSalesExpenses.Filter = "DescriptionID = " & toNumber(.TextMatrix(c, 3))
                    
            rsSalesExpenses![Amount] = .TextMatrix(c, 2)

            rsSalesExpenses.Update
        Next c
    End With

    'Clear variables
    c = 0
    Set rsSalesExpenses = Nothing
    
    CN.CommitTrans
    
    Unload Me
    
    Exit Sub

err:
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub cmdUpdate_Click()
    If txtAmount.Text < 0 Then txtAmount.SetFocus
    
    Dim CurrRow As Integer
        
    CurrRow = getFlexPos(Grid, 3, txtDescription.Tag)

    'Add to grid
    With Grid

        .Row = CurrRow
        
        .TextMatrix(CurrRow, 1) = txtDescription.Text
        .TextMatrix(CurrRow, 2) = txtAmount.Text
        
        'Highlight the current row's column
        .ColSel = 2
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub Form_Load()
    Dim I As Integer
    Dim intRecordCount As Integer
    
    InitGrid
    
    intRecordCount = getRecordCount("Sales_Expenses", "WHERE ReceiptBatchID=" & PK)
    
    If intRecordCount = 0 Then
        For I = 1 To 4
            CN.Execute "INSERT INTO Sales_Expenses ( ReceiptBatchID, [Date], DescriptionID ) " _
                    & "VALUES (" & PK & "," & Date & "," & I & ")"
        Next I
        
        DisplayForEditing
    Else
        DisplayForEditing
    End If
End Sub
    
Private Sub Grid_Click()
    With Grid
        txtDescription.Text = .TextMatrix(.RowSel, 1)
        txtDescription.Tag = .TextMatrix(.RowSel, 3) 'Add tag coz boundtext is empty
        txtAmount.Text = .TextMatrix(.RowSel, 2)
    End With
End Sub
