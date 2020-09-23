VERSION 5.00
Begin VB.Form frmRoutesAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Area"
   ClientHeight    =   2460
   ClientLeft      =   6345
   ClientTop       =   6870
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAddCharges 
      Height          =   225
      Left            =   1680
      TabIndex        =   10
      Top             =   1380
      Width           =   165
   End
   Begin VB.ComboBox cboArea 
      Height          =   315
      ItemData        =   "frmRoutesAE.frx":0000
      Left            =   1680
      List            =   "frmRoutesAE.frx":000D
      TabIndex        =   8
      Text            =   "North"
      Top             =   990
      Width           =   1485
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2220
      TabIndex        =   3
      Top             =   1980
      Width           =   1125
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3450
      TabIndex        =   4
      Top             =   1980
      Width           =   1125
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   270
      TabIndex        =   2
      Top             =   1980
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1680
      MaxLength       =   3
      TabIndex        =   0
      Top             =   270
      Width           =   555
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   1
      Top             =   630
      Width           =   2805
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   120
      TabIndex        =   5
      Top             =   1830
      Width           =   4635
      _ExtentX        =   10081
      _ExtentY        =   53
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Additional Charges"
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   1380
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Area"
      Height          =   285
      Left            =   510
      TabIndex        =   9
      Top             =   990
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   255
      Left            =   510
      TabIndex        =   7
      Top             =   270
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Description"
      Height          =   255
      Left            =   510
      TabIndex        =   6
      Top             =   630
      Width           =   1095
   End
End
Attribute VB_Name = "frmRoutesAE"
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
Dim RS                      As New Recordset
Dim blnRemarks As Boolean

Private Sub DisplayForEditing()
    On Error GoTo errHandler
    
    With RS
        txtEntry(0).Text = .Fields("Route")
        txtEntry(1).Text = .Fields("Desc")
        cboArea.Text = .Fields("Area")
        chkAddCharges.Value = IIf(.Fields("AddCharges") = True, 1, 0)
    End With

    Exit Sub
errHandler:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).Text = ""
    txtEntry(1).Text = ""
    
    txtEntry(0).SetFocus
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    If Trim(txtEntry(0).Text) = "" Then Exit Sub
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        RS.AddNew
        RS.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        RS.Fields("DateModified") = Now
        RS.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With RS
        .Fields("Route") = txtEntry(0).Text
        .Fields("Desc") = txtEntry(1).Text
        .Fields("Area") = cboArea.Text
        .Fields("AddCharges") = chkAddCharges.Value
        
        .Update
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
err:
'  If err.Number = -2147217887 Then Resume Next
  MsgBox "Error: " & err.Description & vbCr _
  & "Form: frmRoutesAE" & vbCr _
  & "Sub: cmdSave_Click", vbExclamation
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And blnRemarks = False Then
        SendKeys "{TAB}"
    End If
End Sub

Private Sub Form_Load()
    RS.CursorLocation = adUseClient
    RS.Open "SELECT * FROM Routes WHERE RouteID = " & PK, CN, adOpenStatic, adLockOptimistic
        
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        'cmdPH.Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmRoutes.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            'srcTextAdd.Text = rs![DisplayAddr]
            'srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmRoutesAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub





