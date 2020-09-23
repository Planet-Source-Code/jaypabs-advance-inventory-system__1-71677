VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmLocationsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Routes"
   ClientHeight    =   4905
   ClientLeft      =   2235
   ClientTop       =   2640
   ClientWidth     =   5565
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   315
      Left            =   3810
      TabIndex        =   3
      Top             =   1410
      Width           =   495
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2310
      TabIndex        =   6
      Top             =   4410
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3720
      TabIndex        =   7
      Top             =   4410
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   315
      Index           =   1
      Left            =   1140
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   600
      Width           =   1635
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   240
      TabIndex        =   5
      Top             =   4410
      Width           =   1680
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   210
      TabIndex        =   8
      Top             =   4305
      Width           =   5115
      _ExtentX        =   18389
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcRoute 
      Height          =   315
      Left            =   1140
      TabIndex        =   0
      Top             =   240
      Width           =   2745
      _ExtentX        =   4842
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin MSComctlLib.ListView lvCities 
      Height          =   2235
      Left            =   1110
      TabIndex        =   4
      Top             =   1770
      Width           =   3225
      _ExtentX        =   5689
      _ExtentY        =   3942
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
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
         Text            =   "Cities"
         Object.Width           =   4683
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "City ID"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "LocationID"
         Object.Width           =   0
      EndProperty
   End
   Begin MSDataListLib.DataCombo dcCities 
      Height          =   315
      Left            =   1110
      TabIndex        =   2
      Top             =   1410
      Width           =   2655
      _ExtentX        =   4683
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
   Begin VB.Label Label2 
      Caption         =   "Cites"
      Height          =   285
      Left            =   1140
      TabIndex        =   11
      Top             =   1140
      Width           =   705
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Desc"
      Height          =   240
      Index           =   0
      Left            =   330
      TabIndex        =   10
      Top             =   600
      Width           =   675
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   240
      Index           =   11
      Left            =   300
      TabIndex        =   9
      Top             =   240
      Width           =   735
   End
End
Attribute VB_Name = "frmLocationsAE"
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

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Dim rsPMeasures             As New Recordset
Dim d()                     As Long     'this array will hold the id of those productunit(s) that will be deleted
Dim b                       As Long   'bookmark to the next element on the arrary d
Dim pm_ID                   As Long   'position of the unit on the list

Private Sub DisplayForEditing()
    On Error GoTo erR
    
    Dim rs1 As New Recordset
    
    rs1.CursorLocation = adUseClient
    rs1.Open "SELECT * FROM qry_Locations WHERE RouteID = " & PK, CN, adOpenStatic, adLockOptimistic
      
    ReDim d(rs1.RecordCount)      're-define array dimension
    
    With rs1
      dcRoute.Text = .Fields("Route")
      txtEntry(1).Text = .Fields("Desc")
    End With
    
    
    'load units under stock
    With lvCities
      dcCities.Text = ""
      Do While Not rs1.EOF
        .ListItems.Add
        .ListItems(.ListItems.Count).SubItems(1) = rs1!City
        .ListItems(.ListItems.Count).SubItems(2) = rs1!CityID
        .ListItems(.ListItems.Count).SubItems(3) = rs1!LocationID
        
        rs1.MoveNext
      Loop
    End With
    
    
    Exit Sub
erR:
        'If err.Number = 94 Then Resume Next
        MsgBox "Error: " & erR.Description, vbExclamation
End Sub

Private Sub cmdAdd_Click()
  
  With lvCities
    If Trim(dcRoute.Text) = "" Or Trim(dcCities.Text) = "" Then Exit Sub
          
    
      If pm_ID = 0 Then
        If Not AlreadyAdded(dcCities.Text) Then
          .ListItems.Add
          .ListItems(.ListItems.Count).SubItems(1) = dcCities.Text
          .ListItems(.ListItems.Count).SubItems(2) = dcCities.BoundText
        End If
      Else
        .ListItems(pm_ID).SubItems(1) = dcCities.Text
        .ListItems(pm_ID).SubItems(2) = dcCities.BoundText
      End If
            
      dcCities.Text = ""
  End With
End Sub

Private Function AlreadyAdded(ByVal City As String) As Boolean
  Dim I As Long
  With lvCities
    For I = 1 To .ListItems.Count
      If .ListItems(I).SubItems(1) = City Then
        .ListItems(I).SubItems(1) = dcCities.Text
        
        dcCities.Text = ""
        AlreadyAdded = True
        Exit Function
      End If
    Next
    
    AlreadyAdded = False
  End With
End Function



Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    dcRoute.Text = ""
    dcCities.Text = ""
    lvCities.ListItems.Clear
    txtEntry(1) = ""
End Sub


Private Sub CmdSave_Click()
On Error GoTo erR
  
  'check for blank category
  If Trim(dcRoute.Text) = "" Then
    MsgBox "Please specify a route for the cities.", vbExclamation
    Exit Sub
  End If
  'check for blank unit measures
  If lvCities.ListItems.Count < 0 Then
    MsgBox "Please provide at least one city.", vbExclamation
    Exit Sub
  End If
  
      
 ' If State = adStateAddMode Or State = adStatePopupMode Then
  '  rs.AddNew
    'rs.Fields("StockId") = PK
    
  'Else
    
  'End If
  
  'Dim RouteID As Long
  
  'With rs
  '  .Fields("CityID") = GetCityID(lvCities.ListItems(Y).SubItems(1))
  '  .Fields("RouteID") = dcRoute.BoundText
  '
  '  .Update
  'End With
  
  'delete all stockunit on the array d
  Dim ctr As Long
  For ctr = 0 To b - 1
    CN.Execute "Delete * From Locations Where LocationID=" & d(ctr)
  Next
    
  'save stockunit
  'rsPMeasures.CursorLocation = adUseClient
  'rsPMeasures.Open "SELECT * FROM Stock_Unit WHERE StockId = " & PK, CN, adOpenStatic, adLockOptimistic
  
  Dim I As Long
  With lvCities
    
    For I = 1 To .ListItems.Count
      If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
        If rs.State = 1 Then rs.Close
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM Locations", CN, adOpenStatic, adLockOptimistic
        
        'new stock new unit(s)
        rs.AddNew
        'rs.Fields("addedbyfk") = CurrUser.USER_PK
        rs.Fields("CityID") = GetCityID(.ListItems(I).SubItems(1))
        rs.Fields("RouteID") = dcRoute.BoundText
        rs.Update
        
      Else
        If LocationExist(dcRoute.Text, .ListItems(I).SubItems(1)) Then
          'update
          If rs.State = 1 Then rs.Close
          rs.CursorLocation = adUseClient
          rs.Open "SELECT * FROM Locations WHERE LocationID = " & .ListItems(I).SubItems(3), CN, adOpenStatic, adLockOptimistic
          
          rs.Fields("CityID") = .ListItems(I).SubItems(2)
          rs.Fields("RouteID") = dcRoute.BoundText
          rs.Update
          rs.MoveNext
          
         ' rs.Fields("datemodified") = Now
         ' rs.Fields("lastuserfk") = CurrUser.USER_PK
        Else
          GoTo AddNew       'if new unit added to the existing unit(s)
        End If
      End If
      
      
      
    Next
  End With
  
  
    
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
  PK = getIndex("Locations")
End Sub


Private Function GetCityID(ByVal City As String) As Long
  Dim sql As String
  Dim rstemp As New Recordset
  
  sql = "SELECT Cities.CityID " _
  & "From Cities WHERE (((Cities.City)='" & Replace(City, "'", "''") & "'))"
  
  rstemp.Open sql, CN, adOpenDynamic, adLockOptimistic
  If Not rstemp.EOF Then GetCityID = rstemp!CityID
  
  Set rstemp = Nothing
End Function





Private Sub Form_Load()
   'If rs.State = 1 Then rs.Close
    
    
    'rs1.CursorLocation = adUseClient
    'rs1.Open "SELECT * FROM qry_Stock_Unit WHERE StockId = " & PK, CN, adOpenStatic, adLockOptimistic
    
    'bind_dc "SELECT * FROM Stocks_Category order by category asc", "Category", dcRoute, "CategoryID", True
    bind_dc "SELECT * FROM Routes Order By Route Asc", "Route", dcRoute, "RouteID", True
    bind_dc "SELECT * FROM Cities Order By City Asc", "City", dcCities, "CityID", True
        
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        dcRoute.Text = ""
        dcCities.Text = ""
        'GeneratePK
    Else
      b = 0
        
        Caption = "Edit Entry"
        DisplayForEditing
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmLocations.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmLocationsAE = Nothing
End Sub

Private Sub lvList_BeforeLabelEdit(Cancel As Integer)

End Sub

Private Sub lvList_DblClick()
  
End Sub



'Private Sub lvPriceHistory_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'  With lvPriceHistory
    'MsgBox .ColumnHeaders(2).Width & vbCr _
    & .ColumnHeaders(3).Width & vbCr _
    & .ColumnHeaders(4).Width
  
'  End With
'End Sub

Private Sub lvCities_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  With lvCities
    'MsgBox .ColumnHeaders(2).Width & vbCr _
    & .ColumnHeaders(3).Width & vbCr _
    & .ColumnHeaders(4).Width & vbCr _
    & .ColumnHeaders(5).Width & vbCr _
    & .ColumnHeaders(6).Width & vbCr _

  End With
End Sub

Private Sub lvCities_DblClick()
  With lvCities
    dcCities.Text = .ListItems(.SelectedItem.Index).SubItems(1)
    'txtEntry(10).Text = .ListItems(.SelectedItem.Index).SubItems(5)
    'dcCities.Text = .ListItems(.SelectedItem.Index).SubItems(2)
    pm_ID = .SelectedItem.Index
    'txtEntry(10).Text = .ListItems(.SelectedItem.Index).SubItems(4)
    'dcChild.Text = .ListItems(.SelectedItem.Index).SubItems(5)
  End With
End Sub

Private Sub lvCities_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 46 Then
    With lvCities
      If .ListItems.Count <= 0 Then Exit Sub
      
      If .ListItems(.SelectedItem.Index).SubItems(3) <> "" Then
        d(b) = .ListItems(.SelectedItem.Index).SubItems(3)        'place unitid to the element b of the array
        b = b + 1
      End If
      
      .ListItems.Remove (.SelectedItem.Index)
    End With
  End If
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






