VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLocalForwarderAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Local Forwarder"
   ClientHeight    =   4890
   ClientLeft      =   2235
   ClientTop       =   3150
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   450
      Picture         =   "frmLocalForwarderAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Remove"
      Top             =   2820
      Visible         =   0   'False
      Width           =   275
   End
   Begin MSDataListLib.DataCombo dcTitle 
      Height          =   315
      Left            =   390
      TabIndex        =   5
      Top             =   2340
      Width           =   3165
      _ExtentX        =   5583
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   0
      Top             =   270
      Width           =   4155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2910
      TabIndex        =   9
      Top             =   4380
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4320
      TabIndex        =   10
      Top             =   4380
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "&Modification History"
      Height          =   315
      Left            =   270
      TabIndex        =   8
      Top             =   4380
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   1
      Top             =   600
      Width           =   3735
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   2
      Top             =   930
      Width           =   2775
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1260
      Width           =   1215
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1440
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1590
      Width           =   2265
   End
   Begin VB.TextBox txtAmount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3690
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2340
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Update"
      Height          =   315
      Left            =   4950
      TabIndex        =   7
      Top             =   2370
      Width           =   720
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   180
      TabIndex        =   11
      Top             =   4230
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1455
      Left            =   390
      TabIndex        =   20
      Top             =   2700
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2566
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
   Begin VB.Label Label1 
      Caption         =   "Company"
      Height          =   285
      Left            =   360
      TabIndex        =   18
      Top             =   300
      Width           =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "Area "
      Height          =   285
      Left            =   360
      TabIndex        =   17
      Top             =   630
      Width           =   1005
   End
   Begin VB.Label Label3 
      Caption         =   "City "
      Height          =   285
      Left            =   360
      TabIndex        =   16
      Top             =   960
      Width           =   1005
   End
   Begin VB.Label Label4 
      Caption         =   "Tel. No."
      Height          =   285
      Left            =   360
      TabIndex        =   15
      Top             =   1290
      Width           =   1005
   End
   Begin VB.Label Label5 
      Caption         =   "Mobile"
      Height          =   285
      Left            =   360
      TabIndex        =   14
      Top             =   1620
      Width           =   1005
   End
   Begin VB.Label Label6 
      Caption         =   "Title"
      Height          =   285
      Left            =   390
      TabIndex        =   13
      Top             =   2100
      Width           =   915
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Amount"
      Height          =   285
      Left            =   3690
      TabIndex        =   12
      Top             =   2100
      Width           =   1185
   End
End
Attribute VB_Name = "frmLocalForwarderAE"
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
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo erR
    
    With rs
      txtEntry(0).Text = .Fields("LocalForwarder")
      txtEntry(1).Text = .Fields("Area")
      txtEntry(2).Text = .Fields("City")
      txtEntry(3).Text = .Fields("Telephone")
      txtEntry(4).Text = .Fields("Mobile")
    End With
    
    'Display the details
    Dim RSDetails As New Recordset

    cIRowCount = 0
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Local_Forwarder_Details WHERE LocalForwarderID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 3) = "" Then
                    .TextMatrix(1, 1) = RSDetails![AccTitle]
                    .TextMatrix(1, 2) = RSDetails![Amount]
                    .TextMatrix(1, 3) = RSDetails![AccountDescriptionID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![AccTitle]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Amount]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![AccountDescriptionID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 3
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing
    
    Exit Sub
erR:
    If erR.Number = 94 Then Resume Next
    
    MsgBox "Error: " & erR.Description, vbExclamation
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

Private Sub cmdAdd_Click()
    If dcTitle.Text = "" Or txtAmount.Text = "" Then dcTitle.SetFocus: Exit Sub

    Dim CurrRow As Integer
    Dim intDetailID As Integer
    
    CurrRow = getFlexPos(Grid, 3, dcTitle.BoundText)
    intDetailID = dcTitle.BoundText

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 3) = "" Then
                .TextMatrix(1, 1) = dcTitle.Text
                .TextMatrix(1, 2) = txtAmount.Text
                .TextMatrix(1, 3) = intDetailID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = dcTitle.Text
                .TextMatrix(.Rows - 1, 2) = txtAmount.Text
                .TextMatrix(.Rows - 1, 3) = intDetailID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = dcTitle.Text
                .TextMatrix(CurrRow, 2) = txtAmount.Text
                .TextMatrix(CurrRow, 3) = intDetailID
            Else
                Exit Sub
            End If
        End If
        
        'Highlight the current row's column
        .ColSel = 3
        'Display a remove button
        Grid_Click
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
On Error GoTo erR

    'check for blank product
    If Trim(txtEntry(0).Text) = "" Then
        MsgBox "Company should not be empty.", vbExclamation
        Exit Sub
    End If
    
    If cIRowCount < 1 Then
        MsgBox "Please provide at least one title for local services.", vbExclamation
        Exit Sub
    End If
    
    CN.BeginTrans
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("LocalForwarderID") = PK
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("LocalForwarder") = txtEntry(0).Text
        .Fields("Area") = txtEntry(1).Text
        .Fields("City") = txtEntry(2).Text
        .Fields("Telephone") = txtEntry(3).Text
        .Fields("Mobile") = txtEntry(4).Text
        
        .Update
    End With
    
    Dim RSDetails As New Recordset
       
    'save stockunit
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Local_Forwarder_Details WHERE LocalForwarderID = " & PK, CN, adOpenStatic, adLockOptimistic
  
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSDetails.AddNew

                RSDetails![LocalForwarderID] = PK
                RSDetails![AccountDescriptionID] = toNumber(.TextMatrix(c, 3))
                RSDetails![Amount] = toNumber(.TextMatrix(c, 2))

                RSDetails.Update
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "AccountDescriptionID = " & toNumber(.TextMatrix(c, 3))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew

                RSDetails![LocalForwarderID] = PK
                RSDetails![AccountDescriptionID] = toNumber(.TextMatrix(c, 3))
                RSDetails![Amount] = toNumber(.TextMatrix(c, 2))

                RSDetails.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set RSDetails = Nothing
    
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
    MsgBox "Error: " & erR.Description, vbExclamation
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

'Procedure used to generate PK
Private Sub GeneratePK()
  PK = getIndex("Local_Forwarder")
End Sub

Private Sub Form_Load()
    InitGrid
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Local_Forwarder WHERE LocalForwarderID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Local_Forwarder_Account_Description", "AccTitle", dcTitle, "LocalForwarderAccTitleID", True
        
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

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmLocalForwarder.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
        End If
    End If
    
    Set frmLocalForwarderAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        dcTitle.Text = .TextMatrix(.RowSel, 1)
        txtAmount.Text = .TextMatrix(.RowSel, 2)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 3) = "" Then '10 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 6 Then KeyAscii = isNumber(KeyAscii)
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
        .ColWidth(1) = 2400
        .ColWidth(2) = 1500
        .ColWidth(3) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Class"
        .TextMatrix(0, 2) = "Cost Handling"
        .TextMatrix(0, 3) = "LocalForwarderDetailID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSDetails As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Local_Forwarder_Detail WHERE LocalForwarderID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            CurrRow = getFlexPos(Grid, 3, RSDetails!AccountDescriptionID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Local_Forwarder_Detail", "LocalForwarderDetailID", "", True, RSDetails!LocalForwarderDetailID
                End If
            End With
            RSDetails.MoveNext
        Wend
    End If
End Sub
