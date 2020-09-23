VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmCustomersAE 
   BorderStyle     =   0  'None
   Caption         =   "Edit Entry"
   ClientHeight    =   6465
   ClientLeft      =   1410
   ClientTop       =   2460
   ClientWidth     =   11070
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6465
   ScaleWidth      =   11070
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   14
      Left            =   1620
      TabIndex        =   10
      Top             =   4350
      Width           =   1035
   End
   Begin VB.TextBox txtEntry 
      Alignment       =   1  'Right Justify
      Height          =   285
      Index           =   15
      Left            =   1620
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   4710
      Width           =   1035
   End
   Begin VB.CheckBox chkBlackListed 
      Alignment       =   1  'Right Justify
      Caption         =   "Black Listed"
      Height          =   195
      Left            =   570
      TabIndex        =   12
      Top             =   5070
      Width           =   1245
   End
   Begin VB.CommandButton cmdPH 
      Caption         =   "Purchase History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2310
      TabIndex        =   21
      Top             =   5865
      Width           =   1590
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   510
      TabIndex        =   20
      Top             =   5865
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   8
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3645
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   7
      Top             =   3315
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1905
      Width           =   2505
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1620
      MaxLength       =   200
      TabIndex        =   2
      Top             =   1530
      Width           =   2505
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   795
      Width           =   2505
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   9360
      TabIndex        =   23
      Top             =   5880
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   7890
      TabIndex        =   22
      Top             =   5880
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1620
      TabIndex        =   4
      Top             =   2280
      Width           =   2505
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   9
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3990
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1620
      TabIndex        =   6
      Top             =   2970
      Width           =   2505
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bank Infos"
      Height          =   3315
      Left            =   4260
      TabIndex        =   24
      Top             =   720
      Width           =   6525
      Begin VB.CommandButton btnRemove 
         Height          =   275
         Left            =   180
         Picture         =   "frmCustomersAE.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Remove"
         Top             =   1110
         Visible         =   0   'False
         Width           =   275
      End
      Begin VB.TextBox txtBranch 
         Height          =   285
         Left            =   1620
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   540
         Width           =   1575
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Height          =   315
         Left            =   5670
         TabIndex        =   17
         Top             =   540
         Width           =   765
      End
      Begin VB.TextBox txtAcctName 
         Height          =   285
         Left            =   4290
         TabIndex        =   16
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox txtAcctNo 
         Height          =   285
         Left            =   3270
         TabIndex        =   15
         Top             =   540
         Width           =   975
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
         Height          =   2190
         Left            =   120
         TabIndex        =   44
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
      Begin ctrlNSDataCombo.NSDataCombo nsdBank 
         Height          =   315
         Left            =   120
         TabIndex        =   13
         Top             =   540
         Width           =   1470
         _ExtentX        =   2593
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
         Left            =   150
         TabIndex        =   39
         Top             =   300
         Width           =   525
      End
      Begin VB.Label Label4 
         Caption         =   "Account Name"
         Height          =   435
         Left            =   4290
         TabIndex        =   27
         Top             =   120
         Width           =   705
      End
      Begin VB.Label Label3 
         Caption         =   "Acct. No."
         Height          =   225
         Left            =   3270
         TabIndex        =   26
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Branch"
         Height          =   225
         Left            =   1650
         TabIndex        =   25
         Top             =   300
         Width           =   525
      End
   End
   Begin VB.TextBox txtEntry 
      Height          =   1065
      Index           =   16
      Left            =   4260
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   4440
      Width           =   3855
   End
   Begin MSDataListLib.DataCombo dcCity 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   2610
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   300
      TabIndex        =   28
      Top             =   5730
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   1125
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer"
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
      TabIndex        =   43
      Top             =   180
      Width           =   4905
   End
   Begin VB.Shape Shape1 
      Height          =   5715
      Left            =   150
      Top             =   630
      Width           =   10725
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Term"
      Height          =   240
      Left            =   240
      TabIndex        =   42
      Top             =   4350
      Width           =   1245
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Limit"
      Height          =   240
      Left            =   240
      TabIndex        =   41
      Top             =   4710
      Width           =   1245
   End
   Begin VB.Label Label7 
      Caption         =   "Remarks"
      Height          =   225
      Left            =   4260
      TabIndex        =   40
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   240
      TabIndex        =   38
      Top             =   1125
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Landline"
      Height          =   240
      Index           =   7
      Left            =   240
      TabIndex        =   37
      Top             =   3645
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile"
      Height          =   240
      Index           =   5
      Left            =   240
      TabIndex        =   36
      Top             =   3315
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner's Name"
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   35
      Top             =   1905
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "TIN"
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   34
      Top             =   1530
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Store Name"
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   33
      Top             =   795
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Area Address"
      Height          =   240
      Index           =   12
      Left            =   240
      TabIndex        =   32
      Top             =   2280
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   240
      Index           =   9
      Left            =   240
      TabIndex        =   31
      Top             =   3990
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City"
      Height          =   240
      Index           =   16
      Left            =   240
      TabIndex        =   30
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchaser Name"
      Height          =   240
      Index           =   17
      Left            =   240
      TabIndex        =   29
      Top             =   2970
      Width           =   1245
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Solid
      Height          =   435
      Left            =   150
      Top             =   150
      Width           =   10725
   End
   Begin VB.Shape Shape3 
      BorderWidth     =   2
      Height          =   6375
      Left            =   60
      Top             =   60
      Width           =   10935
   End
End
Attribute VB_Name = "frmCustomersAE"
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
Dim rsClientBank As New Recordset

Private Sub btnAdd_Click()
    If nsdBank.Text = "" Or txtAcctNo.Text = "" Or txtAcctName.Text = "" Then nsdBank.SetFocus: Exit Sub

    Dim CurrRow As Integer
    Dim intBankID As Integer
    
    If nsdBank.BoundText = "" Then
        CurrRow = getFlexPos(Grid, 5, nsdBank.Tag)
        intBankID = nsdBank.Tag
    Else
        CurrRow = getFlexPos(Grid, 5, nsdBank.BoundText)
        intBankID = nsdBank.BoundText
    End If

    'Add to grid
    With Grid
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
            cIRowCount = cIRowCount + 1
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
        Grid_Click
    End With
End Sub

Private Sub DisplayForEditing()
    On Error GoTo err
    Dim rsClients As New Recordset
    
    rsClients.CursorLocation = adUseClient
    rsClients.Open "SELECT * FROM qry_Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    With rsClients
        txtEntry(1).Text = .Fields("Company")
        dcCategory.BoundText = .Fields![CategoryID]
        txtEntry(2).Text = .Fields("Tin")
        txtEntry(3).Text = .Fields("OwnersName")
        txtEntry(4).Text = .Fields("Address")
        dcCity.BoundText = .Fields![CityID]
        txtEntry(6).Text = .Fields("PurchaserName")
        txtEntry(7).Text = .Fields("Mobile")
        txtEntry(8).Text = .Fields("Landline")
        txtEntry(9).Text = .Fields("Fax")
        txtEntry(14).Text = .Fields("CreditTerm")
        txtEntry(15).Text = .Fields("CreditLimit")
        chkBlackListed.Value = IIf(.Fields("BlackListed") = True, 1, 0)
        txtEntry(16).Text = .Fields("Remarks")
    End With
    
    'Display the details
    Dim rsClientBank As New Recordset

    cIRowCount = 0
    
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    If rsClientBank.RecordCount > 0 Then
        rsClientBank.MoveFirst
        While Not rsClientBank.EOF
          cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 5) = "" Then
                    .TextMatrix(1, 1) = rsClientBank![Bank]
                    .TextMatrix(1, 2) = rsClientBank![Branch]
                    .TextMatrix(1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(1, 4) = rsClientBank![AccountName]
                    .TextMatrix(1, 5) = rsClientBank![BankID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rsClientBank![Bank]
                    .TextMatrix(.Rows - 1, 2) = rsClientBank![Branch]
                    .TextMatrix(.Rows - 1, 3) = rsClientBank![AccountNo]
                    .TextMatrix(.Rows - 1, 4) = rsClientBank![AccountName]
                    .TextMatrix(.Rows - 1, 5) = rsClientBank![BankID]
                End If
            End With
            rsClientBank.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 5
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
            Grid.FixedCols = 1
        End If
    End If

    rsClientBank.Close
    'Clear variables
    Set rsClientBank = Nothing
        
    'txtEntry(1).SetFocus
    Exit Sub
err:
    If err.Number = 94 Then Resume Next
    
    prompt_err err, Name, "DisplayForEditing"
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
  
  txtEntry(15).Text = "0.00"
  txtEntry(1).SetFocus
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewer.CUS_PK = PK
    'frmInvoiceViewer.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.lblTitle.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.show vbModal
End Sub

Private Sub cmdSave_Click()
    On Error GoTo err

    If Trim(txtEntry(1).Text) = "" Then Exit Sub
    
    CN.BeginTrans

    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        
        rs.Fields("ClientID") = PK
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
      .Fields("Company") = txtEntry(1).Text
      .Fields("CategoryID") = dcCategory.BoundText
      .Fields("Tin") = txtEntry(2).Text
      .Fields("OwnersName") = txtEntry(3).Text
      .Fields("Address") = txtEntry(4).Text
      .Fields("CityID") = dcCity.BoundText
      .Fields("PurchaserName") = txtEntry(6).Text
      .Fields("Mobile") = txtEntry(7).Text
      .Fields("Landline") = txtEntry(8).Text
      .Fields("Fax") = txtEntry(9).Text
      .Fields("CreditTerm") = toNumber(txtEntry(14).Text)
      .Fields("CreditLimit") = toNumber(txtEntry(15).Text)
      .Fields("BlackListed") = IIf(chkBlackListed.Value = 1, True, False)
      .Fields("Remarks") = txtEntry(16).Text
       
      .Update
    End With

    Dim rsClientBank As New Recordset

    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    
    DeleteItems
    
    Dim c As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                rsClientBank.AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            ElseIf State = adStateEditMode Then
                rsClientBank.Filter = "BankID = " & toNumber(.TextMatrix(c, 5))
            
                If rsClientBank.RecordCount = 0 Then GoTo AddNew

                rsClientBank![ClientID] = PK
                rsClientBank![BankID] = toNumber(.TextMatrix(c, 5))
                rsClientBank![AccountNo] = .TextMatrix(c, 3)
                rsClientBank![AccountName] = .TextMatrix(c, 4)

                rsClientBank.Update
            End If

        Next c
    End With

    'Clear variables
    c = 0
    Set rsClientBank = Nothing
    
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
    CN.RollbackTrans
    prompt_err err, Name, "cmdSave_Click"
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
    rs.Open "SELECT * FROM Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
        
    rsClientBank.CursorLocation = adUseClient
    rsClientBank.Open "SELECT * FROM qry_Clients_Bank WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Clients_Category", "Category", dcCategory, "CategoryID", True
    bind_dc "SELECT * FROM Cities", "City", dcCity, "CityID", True
   
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        GeneratePK
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        cmdPH.Enabled = True
    End If

End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Clients")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmCustomers.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmCustomersAE = Nothing
End Sub

Private Sub ResetEntry()
    txtBranch.Text = ""
    txtAcctNo.Text = ""
    txtAcctName.Text = ""
End Sub

Private Sub Grid_Click()
    With Grid
        nsdBank.Text = .TextMatrix(.RowSel, 1)
        nsdBank.Tag = .TextMatrix(.RowSel, 5) 'Add tag coz boundtext is empty
        txtBranch.Text = .TextMatrix(.RowSel, 2)
        txtAcctNo.Text = .TextMatrix(.RowSel, 3)
        txtAcctName.Text = .TextMatrix(.RowSel, 4)
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 5) = "" Then '10 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub nsdBank_Change()
    txtBranch.Text = nsdBank.getSelValueAt(3)
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 16 Then
        blnRemarks = True
        Exit Sub
    Else
        blnRemarks = False
    End If
    
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_KeyPress(Index As Integer, KeyAscii As Integer)
  If Index = 15 Then KeyAscii = isNumber(KeyAscii)
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
End Sub

Private Sub InitNSD()
    'For Bank
    With nsdBank
        .ClearColumn
        .AddColumn "Bank ID", 1794.89
        .AddColumn "Bank", 2264.88
        .AddColumn "Branch", 2670.23
        .Connection = CN.ConnectionString
        
        '.sqlFields = "VendorID, Company, Location"
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

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSBank As New Recordset
    
    If State = adStateAddMode Then Exit Sub
    
    RSBank.CursorLocation = adUseClient
    RSBank.Open "SELECT * FROM Clients_Bank WHERE ClientID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSBank.RecordCount > 0 Then
        RSBank.MoveFirst
        While Not RSBank.EOF
            CurrRow = getFlexPos(Grid, 5, RSBank!BankID)
        
            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Clients_Bank", "PK", "", True, RSBank!PK
                End If
            End With
            RSBank.MoveNext
        Wend
    End If
End Sub

Private Sub txtEntry_Validate(Index As Integer, Cancel As Boolean)
    txtEntry(14).Text = toNumber(txtEntry(14).Text)
End Sub
