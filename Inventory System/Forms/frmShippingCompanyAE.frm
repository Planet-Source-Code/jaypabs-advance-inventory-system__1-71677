VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmShippingCompanyAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Forwarders"
   ClientHeight    =   5595
   ClientLeft      =   1275
   ClientTop       =   2445
   ClientWidth     =   7875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   7875
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   480
      Picture         =   "frmShippingCompanyAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Remove"
      Top             =   3450
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.Frame Frame2 
      Caption         =   "Cargo"
      Height          =   1365
      Left            =   3960
      TabIndex        =   21
      Top             =   1230
      Width           =   3585
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   8
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   8
         Top             =   930
         Width           =   1935
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   6
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   6
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   7
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile"
         Height          =   285
         Index           =   3
         Left            =   270
         TabIndex        =   24
         Top             =   960
         Width           =   1515
      End
      Begin VB.Label Label8 
         Caption         =   "Tel. No."
         Height          =   285
         Left            =   270
         TabIndex        =   23
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Contact Person"
         Height          =   285
         Index           =   2
         Left            =   270
         TabIndex        =   22
         Top             =   630
         Width           =   1515
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Office"
      Height          =   1365
      Left            =   270
      TabIndex        =   17
      Top             =   1230
      Width           =   3585
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   4
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   4
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   3
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   3
         Top             =   270
         Width           =   1935
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   5
         Left            =   1500
         MaxLength       =   100
         TabIndex        =   5
         Top             =   900
         Width           =   1935
      End
      Begin VB.Label Label5 
         Caption         =   "Contact Person"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   20
         Top             =   630
         Width           =   1515
      End
      Begin VB.Label Label4 
         Caption         =   "Tel. No."
         Height          =   285
         Left            =   270
         TabIndex        =   19
         Top             =   300
         Width           =   1515
      End
      Begin VB.Label Label5 
         Caption         =   "Mobile"
         Height          =   285
         Index           =   0
         Left            =   270
         TabIndex        =   18
         Top             =   960
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add"
      Height          =   315
      Left            =   4980
      TabIndex        =   10
      Top             =   3000
      Width           =   720
   End
   Begin VB.TextBox txtFreight 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   3720
      MaxLength       =   100
      TabIndex        =   9
      Top             =   2970
      Width           =   1215
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   2
      Top             =   840
      Width           =   6165
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   1
      Top             =   510
      Width           =   6165
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "&Modification History"
      Height          =   315
      Left            =   210
      TabIndex        =   11
      Top             =   5100
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6240
      TabIndex        =   13
      Top             =   5100
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   4830
      TabIndex        =   12
      Top             =   5100
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   0
      Top             =   180
      Width           =   4155
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   14
      Top             =   4950
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   53
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   1530
      Left            =   420
      TabIndex        =   29
      Top             =   3330
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   2699
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
   Begin ctrlNSDataCombo.NSDataCombo nsdCargo 
      Height          =   315
      Left            =   420
      TabIndex        =   30
      Top             =   2970
      Width           =   3180
      _ExtentX        =   5609
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
      Caption         =   "Company"
      Height          =   285
      Left            =   270
      TabIndex        =   27
      Top             =   210
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Area "
      Height          =   285
      Left            =   270
      TabIndex        =   26
      Top             =   540
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "City "
      Height          =   285
      Left            =   270
      TabIndex        =   25
      Top             =   870
      Width           =   1035
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Cost"
      Height          =   195
      Left            =   3720
      TabIndex        =   16
      Top             =   2730
      Width           =   1185
   End
   Begin VB.Label Label6 
      Caption         =   "Classification"
      Height          =   225
      Left            =   420
      TabIndex        =   15
      Top             =   2700
      Width           =   915
   End
End
Attribute VB_Name = "frmShippingCompanyAE"
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
Dim rs1                     As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo erR
    
    With rs
      txtEntry(0).Text = .Fields("ShippingCompany")
      txtEntry(1).Text = .Fields("Area")
      txtEntry(2).Text = .Fields("City")
      txtEntry(3).Text = .Fields("Telephone")
      txtEntry(4).Text = .Fields("ContactPerson1")
      txtEntry(5).Text = .Fields("Mobile")
      txtEntry(6).Text = .Fields("Telephone2")
      txtEntry(7).Text = .Fields("ContactPerson2")
      txtEntry(8).Text = .Fields("Mobile2")
      
    End With
    
    'Display the details
    Dim rsCargoClass As New Recordset

    cIRowCount = 0
    
    rsCargoClass.CursorLocation = adUseClient
    rsCargoClass.Open "SELECT * FROM qry_Cargo_Class WHERE ShippingCompanyID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsCargoClass.RecordCount > 0 Then
        rsCargoClass.MoveFirst
        While Not rsCargoClass.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 3) = "" Then
                    .TextMatrix(1, 1) = rsCargoClass![Cargo]
                    .TextMatrix(1, 2) = rsCargoClass![Freight]
                    .TextMatrix(1, 3) = rsCargoClass![CargoID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsCargoClass![Cargo]
                    .TextMatrix(.Rows - 1, 2) = rsCargoClass![Freight]
                    .TextMatrix(.Rows - 1, 3) = rsCargoClass![CargoID]
                End If
            End With
            rsCargoClass.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 3
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsCargoClass.Close
    'Clear variables
    Set rsCargoClass = Nothing
    
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
    If nsdCargo.Text = "" Or txtFreight.Text = "" Then nsdCargo.SetFocus: Exit Sub

    Dim CurrRow As Integer
    Dim intCargoID As Integer
    
    If nsdCargo.BoundText = "" Then
        CurrRow = getFlexPos(Grid, 3, nsdCargo.Tag)
        intCargoID = nsdCargo.Tag
    Else
        CurrRow = getFlexPos(Grid, 3, nsdCargo.BoundText)
        intCargoID = nsdCargo.BoundText
    End If

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 3) = "" Then
                .TextMatrix(1, 1) = nsdCargo.Text
                .TextMatrix(1, 2) = txtFreight.Text
                .TextMatrix(1, 3) = intCargoID
            Else
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdCargo.Text
                .TextMatrix(.Rows - 1, 2) = txtFreight.Text
                .TextMatrix(.Rows - 1, 3) = intCargoID

                .Row = .Rows - 1
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already added. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                
                .TextMatrix(CurrRow, 1) = nsdCargo.Text
                .TextMatrix(CurrRow, 2) = txtFreight.Text
                .TextMatrix(CurrRow, 3) = intCargoID
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
        MsgBox "Shipping company should not be empty.", vbExclamation
        Exit Sub
    End If
       
    CN.BeginTrans
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("ShippingCompanyID") = PK
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("ShippingCompany") = txtEntry(0).Text
        .Fields("Area") = txtEntry(1).Text
        .Fields("City") = txtEntry(2).Text
        .Fields("Telephone") = txtEntry(3).Text
        .Fields("ContactPerson1") = txtEntry(4).Text
        .Fields("Mobile") = txtEntry(5).Text
        .Fields("Telephone2") = txtEntry(6).Text
        .Fields("ContactPerson2") = txtEntry(7).Text
        .Fields("Mobile2") = txtEntry(8).Text
        
        .Update
    End With
       
    Dim rsCargoClass As New Recordset
       
    'save stockunit
    rsCargoClass.CursorLocation = adUseClient
    rsCargoClass.Open "SELECT * FROM Cargo_Class WHERE ShippingCompanyID = " & PK, CN, adOpenStatic, adLockOptimistic
  
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                rsCargoClass.AddNew

                rsCargoClass![ShippingCompanyID] = PK
                rsCargoClass![CargoID] = toNumber(.TextMatrix(c, 3))
                rsCargoClass![Freight] = .TextMatrix(c, 2)

                rsCargoClass.Update
            ElseIf State = adStateEditMode Then
                rsCargoClass.Filter = "CargoID = " & toNumber(.TextMatrix(c, 3))
            
                If rsCargoClass.RecordCount = 0 Then GoTo AddNew

                rsCargoClass![ShippingCompanyID] = PK
                rsCargoClass![CargoID] = toNumber(.TextMatrix(c, 3))
                rsCargoClass![Freight] = .TextMatrix(c, 2)

                rsCargoClass.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set rsCargoClass = Nothing
    
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

'Procedure used to generate PK
Private Sub GeneratePK()
  PK = getIndex("Shipping_Company")
End Sub

Private Sub Form_Load()
    InitGrid
    InitNSD
    
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Shipping_Company WHERE ShippingCompanyID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    rs1.CursorLocation = adUseClient
    rs1.Open "SELECT * FROM qry_Cargo_Class WHERE ShippingCompanyID = " & PK, CN, adOpenStatic, adLockOptimistic
        
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        'dcCategory.Text = ""
        'dcParent.Text = ""
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmShippingCompany.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmShippingCompanyAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        nsdCargo.Text = .TextMatrix(.RowSel, 1)
        nsdCargo.Tag = .TextMatrix(.RowSel, 3) 'Add tag coz boundtext is empty
        txtFreight.Text = .TextMatrix(.RowSel, 2)
    
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

Private Sub txtEntry_LostFocus(Index As Integer)
    'If Index = 8 Then cmdSave.Default = True
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
        .TextMatrix(0, 1) = "Cargo"
        .TextMatrix(0, 2) = "Freight"
        .TextMatrix(0, 3) = "CargoID"
        'Set the column alignment
        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
    End With
End Sub

Private Sub InitNSD()
    'For Bank
    With nsdCargo
        .ClearColumn
        .AddColumn "Cargo ID", 1794.89
        .AddColumn "Cargo", 2264.88
        .AddColumn "Loose Cargo", 2670.23
        .Connection = CN.ConnectionString
        
        '.sqlFields = "VendorID, Company, Location"
        .sqlFields = "CargoID, Cargo, LooseCargo"
        .sqlTables = "Cargos"
        .sqlSortOrder = "Cargo ASC"
        
        .BoundField = "CargoID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Cargo Record"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSCargo As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSCargo.CursorLocation = adUseClient
    RSCargo.Open "SELECT * FROM Cargo_Class WHERE ShippingCompanyID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSCargo.RecordCount > 0 Then
        RSCargo.MoveFirst
        While Not RSCargo.EOF
            CurrRow = getFlexPos(Grid, 3, RSCargo!CargoID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Cargo_Class", "CargoClassID", "", True, RSCargo!CargoClassID
                End If
            End With
            RSCargo.MoveNext
        Wend
    End If
End Sub

