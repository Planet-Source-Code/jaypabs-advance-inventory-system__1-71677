VERSION 5.00
Begin VB.Form frmCargosAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cargo"
   ClientHeight    =   1785
   ClientLeft      =   2955
   ClientTop       =   2190
   ClientWidth     =   5040
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoose 
      Alignment       =   1  'Right Justify
      Caption         =   "Loose"
      Height          =   225
      Left            =   780
      TabIndex        =   1
      Top             =   750
      Width           =   915
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Category"
      Top             =   390
      Width           =   3390
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   2160
      TabIndex        =   3
      Top             =   1305
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   3600
      TabIndex        =   4
      Top             =   1305
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   1305
      Width           =   1680
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   105
      TabIndex        =   5
      Top             =   1110
      Width           =   4815
      _ExtentX        =   10213
      _ExtentY        =   53
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cargo Classification"
      Height          =   450
      Left            =   210
      TabIndex        =   6
      Top             =   300
      Width           =   1155
   End
End
Attribute VB_Name = "frmCargosAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public srcText              As TextBox 'Used in pop-up mode

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset

Private Sub DisplayForEditing()
    On Error GoTo erR
    
    With rs
      txtEntry.Text = .Fields("Cargo")
      chkLoose.Value = IIf(.Fields("LooseCargo"), 1, 0)
    End With
    
    Exit Sub
erR:
        If erR.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    chkLoose.Value = 0
    txtEntry.SetFocus
End Sub

Private Sub CmdSave_Click()
    If State = adStateAddMode Then
        rs.AddNew
        rs.Fields("DateAdded") = Now
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    'Phill 2:12
    With rs
        .Fields("Cargo") = txtEntry.Text
        .Fields("LooseCargo") = IIf(chkLoose.Value = 1, True, False)
    
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
    Else
        MsgBox "Changes in  record has been successfully saved.", vbInformation
        Unload Me
    End If
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

Private Sub Form_Load()
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Cargos WHERE CargoID = " & PK, CN, adOpenStatic, adLockOptimistic
    'Check the form state
    If State = adStateAddMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or adStateEditMode Then
            frmCargos.RefreshRecords
        End If
    End If
    
    Set frmCargosAE = Nothing
End Sub





