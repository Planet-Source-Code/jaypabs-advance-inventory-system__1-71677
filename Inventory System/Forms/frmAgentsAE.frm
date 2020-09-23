VERSION 5.00
Begin VB.Form frmAgentsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Agents"
   ClientHeight    =   2955
   ClientLeft      =   2550
   ClientTop       =   3570
   ClientWidth     =   6060
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1800
      TabIndex        =   4
      Top             =   1590
      Width           =   2475
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1800
      TabIndex        =   3
      Top             =   1260
      Width           =   2475
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1800
      TabIndex        =   2
      Top             =   930
      Width           =   1605
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1800
      TabIndex        =   1
      Top             =   600
      Width           =   3885
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3195
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   420
      TabIndex        =   6
      Top             =   2520
      Width           =   1680
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4305
      TabIndex        =   8
      Top             =   2490
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2850
      TabIndex        =   7
      Top             =   2490
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1800
      TabIndex        =   5
      Top             =   1920
      Width           =   1245
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   150
      TabIndex        =   9
      Top             =   2370
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   53
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile"
      Height          =   315
      Left            =   630
      TabIndex        =   15
      Top             =   1560
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Contact No"
      Height          =   315
      Left            =   630
      TabIndex        =   14
      Top             =   1230
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Address"
      Height          =   315
      Left            =   630
      TabIndex        =   13
      Top             =   900
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Code"
      Height          =   315
      Left            =   630
      TabIndex        =   12
      Top             =   570
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Name"
      Height          =   315
      Left            =   630
      TabIndex        =   11
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Commission"
      Height          =   315
      Left            =   630
      TabIndex        =   10
      Top             =   1890
      Width           =   1095
   End
End
Attribute VB_Name = "frmAgentsAE"
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
Dim blnRemarks As Boolean

Private Sub DisplayForEditing()
    On Error GoTo errHandler
    
    With rs
        txtEntry(0).Text = .Fields("AgentName")
        txtEntry(1).Text = .Fields("AgentCode")
        txtEntry(2).Text = .Fields("Address")
        txtEntry(3).Text = .Fields("ContactNo")
        txtEntry(4).Text = .Fields("Mobile")
        txtEntry(5).Text = toNumber(.Fields("Commission"))
    End With
    txtEntry(0).SetFocus
    Exit Sub
errHandler:
        If erR.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).Text = ""
    txtEntry(1).Text = ""
    txtEntry(2).Text = ""
    txtEntry(3).Text = ""
    txtEntry(5).Text = ""
    txtEntry(0).SetFocus
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewer.CUS_PK = PK
    'frmInvoiceViewer.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.lblTitle.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.show vbModal
End Sub

Private Sub CmdSave_Click()
On Error GoTo erR

    If Trim(txtEntry(0).Text) = "" Then Exit Sub
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("AgentName") = txtEntry(0).Text
        .Fields("AgentCode") = txtEntry(1).Text
        .Fields("Address") = txtEntry(2).Text
        .Fields("ContactNo") = txtEntry(3).Text
        .Fields("Mobile") = txtEntry(4).Text
        .Fields("Commission") = toNumber(txtEntry(5).Text)
        
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
erR:
'  If err.Number = -2147217887 Then Resume Next
  MsgBox "Error: " & erR.Description & vbCr _
  & "Form: frmAgentsAE" & vbCr _
  & "Sub: cmdSave_Click", vbExclamation
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
   
    rs.CursorLocation = adUseClient
    rs.Open "SELECT * FROM Agents WHERE AgentID = " & PK, CN, adOpenStatic, adLockOptimistic
        
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
            frmAgents.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            'srcTextAdd.Text = rs![DisplayAddr]
            'srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmAgentsAE = Nothing
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
  If Index = 5 Then KeyAscii = isNumber(KeyAscii)
End Sub
