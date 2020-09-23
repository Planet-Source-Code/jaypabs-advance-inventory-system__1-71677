VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmClientsAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   10095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPH 
      Caption         =   "Purchase History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2160
      TabIndex        =   21
      Top             =   4125
      Width           =   1590
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   360
      TabIndex        =   20
      Top             =   4125
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   8
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   8
      Top             =   3045
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2715
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1305
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1620
      MaxLength       =   200
      TabIndex        =   2
      Top             =   930
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1620
      MaxLength       =   100
      TabIndex        =   0
      Tag             =   "Name"
      Top             =   195
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6975
      TabIndex        =   19
      Top             =   4125
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   5535
      TabIndex        =   18
      Top             =   4125
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1620
      TabIndex        =   4
      Top             =   1680
      Width           =   3855
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   9
      Left            =   1620
      MaxLength       =   20
      TabIndex        =   9
      Top             =   3390
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1620
      TabIndex        =   6
      Top             =   2370
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bank Info"
      Height          =   2265
      Left            =   5670
      TabIndex        =   10
      Top             =   240
      Width           =   4155
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   15
         Left            =   1440
         TabIndex        =   16
         Text            =   "0.00"
         Top             =   1830
         Width           =   2475
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   14
         Left            =   1440
         TabIndex        =   15
         Top             =   1500
         Width           =   2475
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   13
         Left            =   1440
         TabIndex        =   14
         Top             =   1170
         Width           =   2475
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   12
         Left            =   1440
         TabIndex        =   13
         Top             =   840
         Width           =   2475
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   11
         Left            =   1440
         TabIndex        =   12
         Top             =   510
         Width           =   2475
      End
      Begin VB.TextBox txtEntry 
         Height          =   285
         Index           =   10
         Left            =   1440
         TabIndex        =   11
         Top             =   180
         Width           =   2475
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Limit"
         Height          =   315
         Left            =   150
         TabIndex        =   27
         Top             =   1890
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Credit Term"
         Height          =   315
         Left            =   150
         TabIndex        =   26
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Account Name"
         Height          =   315
         Left            =   150
         TabIndex        =   25
         Top             =   1230
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Acct. No."
         Height          =   315
         Left            =   150
         TabIndex        =   24
         Top             =   900
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Branch"
         Height          =   315
         Left            =   150
         TabIndex        =   23
         Top             =   570
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bank Name"
         Height          =   315
         Left            =   150
         TabIndex        =   22
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.TextBox txtEntry 
      Height          =   1185
      Index           =   16
      Left            =   5700
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   2580
      Width           =   4155
   End
   Begin MSDataListLib.DataCombo dcCity 
      Height          =   315
      Left            =   1620
      TabIndex        =   5
      Top             =   2010
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   330
      TabIndex        =   28
      Top             =   4020
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1620
      TabIndex        =   1
      Top             =   525
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   270
      TabIndex        =   38
      Top             =   525
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Landline"
      Height          =   240
      Index           =   7
      Left            =   135
      TabIndex        =   37
      Top             =   3045
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Mobile"
      Height          =   240
      Index           =   5
      Left            =   135
      TabIndex        =   36
      Top             =   2715
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Owner's Name"
      Height          =   240
      Index           =   3
      Left            =   120
      TabIndex        =   35
      Top             =   1305
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "TIN"
      Height          =   240
      Index           =   2
      Left            =   120
      TabIndex        =   34
      Top             =   930
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Store Name"
      Height          =   240
      Index           =   1
      Left            =   270
      TabIndex        =   33
      Top             =   195
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Area Address"
      Height          =   240
      Index           =   12
      Left            =   420
      TabIndex        =   32
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Fax"
      Height          =   240
      Index           =   9
      Left            =   135
      TabIndex        =   31
      Top             =   3390
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "City"
      Height          =   240
      Index           =   16
      Left            =   420
      TabIndex        =   30
      Top             =   2010
      Width           =   1065
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Purchaser Name"
      Height          =   240
      Index           =   17
      Left            =   240
      TabIndex        =   29
      Top             =   2370
      Width           =   1245
   End
End
Attribute VB_Name = "frmClientsAE"
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
    On Error GoTo err
    
    With rs
        txtEntry(1).Text = .Fields("company")
        dcCategory.BoundText = .Fields![CategoryID]
        txtEntry(2).Text = .Fields("tin")
        txtEntry(3).Text = .Fields("ownersname")
        txtEntry(4).Text = .Fields("address")
        dcCity.BoundText = .Fields![CityID]
        txtEntry(6).Text = .Fields("purchasername")
        txtEntry(7).Text = .Fields("mobile")
        txtEntry(8).Text = .Fields("landline")
        txtEntry(9).Text = .Fields("fax")
        txtEntry(10).Text = .Fields("BankName")
        txtEntry(11).Text = .Fields("BankBranch")
        txtEntry(12).Text = .Fields("BankAccountNo")
        txtEntry(13).Text = .Fields("BankAccountName")
        txtEntry(14).Text = .Fields("creditterm")
        txtEntry(15).Text = .Fields("creditlimit")
        txtEntry(16).Text = .Fields("remarks")
    End With
    txtEntry(0).SetFocus
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
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
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("company") = txtEntry(1).Text
        .Fields("categoryid") = dcCategory.BoundText
        .Fields("tin") = txtEntry(2).Text
        .Fields("ownersname") = txtEntry(3).Text
        .Fields("address") = txtEntry(4).Text
        .Fields("CityID") = dcCity.BoundText
        .Fields("purchasername") = txtEntry(6).Text
        .Fields("mobile") = txtEntry(7).Text
        .Fields("landline") = txtEntry(8).Text
        .Fields("fax") = txtEntry(9).Text
        .Fields("BankName") = txtEntry(10).Text
        .Fields("BankBranch") = txtEntry(11).Text
        .Fields("BankAccountNo") = txtEntry(12).Text
        .Fields("BankAccountName") = txtEntry(13).Text
        .Fields("creditterm") = txtEntry(14).Text
        .Fields("creditlimit") = txtEntry(15).Text
        .Fields("remarks") = txtEntry(16).Text
        
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
  & "Form: frmClientsAE" & vbCr _
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
    rs.Open "SELECT * FROM Clients WHERE ClientID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Clients_Category", "Category", dcCategory, "CategoryID", True
    bind_dc "SELECT * FROM Cities", "City", dcCity, "CityID", True
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
    Else
        Caption = "Edit Entry"
        DisplayForEditing
        cmdPH.Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmClients.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmClientsAE = Nothing
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

