VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmStocksAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Edit Entry"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   8865
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   8
      Left            =   1500
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3270
      Width           =   1290
   End
   Begin VB.CommandButton cmdPH 
      Caption         =   "Purchase History"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2010
      TabIndex        =   13
      Top             =   4110
      Width           =   1590
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   210
      TabIndex        =   12
      Top             =   4110
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   5
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   5
      Top             =   2175
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   7
      Left            =   1500
      MaxLength       =   20
      TabIndex        =   7
      Top             =   2910
      Width           =   2490
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   6
      Left            =   1500
      MaxLength       =   10
      TabIndex        =   6
      Top             =   2520
      Width           =   1290
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   4
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   4
      Top             =   1785
      Width           =   2415
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   3
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   3
      Top             =   1410
      Width           =   5055
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1500
      MaxLength       =   200
      TabIndex        =   2
      Top             =   1035
      Width           =   5055
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   1
      Tag             =   "Name"
      Top             =   660
      Width           =   6735
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   6825
      TabIndex        =   11
      Top             =   4110
      Width           =   1335
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   5385
      TabIndex        =   10
      Top             =   4110
      Width           =   1335
   End
   Begin VB.TextBox txtEntry 
      BackColor       =   &H00E6FFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   0
      Left            =   1500
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   285
      Width           =   1965
   End
   Begin InvtySystem.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   180
      TabIndex        =   14
      Top             =   4005
      Width           =   8325
      _ExtentX        =   14684
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCategory 
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Top             =   3630
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "On Hand"
      Height          =   240
      Index           =   8
      Left            =   0
      TabIndex        =   24
      Top             =   3270
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Category"
      Height          =   240
      Index           =   11
      Left            =   150
      TabIndex        =   23
      Top             =   3630
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Reorder Pt."
      Height          =   240
      Index           =   7
      Left            =   0
      TabIndex        =   22
      Top             =   2910
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Sales Price"
      Height          =   240
      Index           =   6
      Left            =   0
      TabIndex        =   21
      Top             =   2175
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Cost"
      Height          =   240
      Index           =   5
      Left            =   0
      TabIndex        =   20
      Top             =   2520
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "ICode"
      Height          =   240
      Index           =   4
      Left            =   0
      TabIndex        =   19
      Top             =   1785
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Short2"
      Height          =   240
      Index           =   3
      Left            =   0
      TabIndex        =   18
      Top             =   1410
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Short1"
      Height          =   240
      Index           =   2
      Left            =   0
      TabIndex        =   17
      Top             =   1035
      Width           =   1365
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Stock"
      Height          =   240
      Index           =   1
      Left            =   150
      TabIndex        =   16
      Top             =   660
      Width           =   1215
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Barcode"
      Height          =   240
      Index           =   0
      Left            =   450
      TabIndex        =   15
      Top             =   285
      Width           =   915
   End
End
Attribute VB_Name = "frmStocksAE"
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

Private Sub DisplayForEditing()
    On Error GoTo err
    
    With rs
        txtEntry(0).Text = .Fields("Barcode")
        txtEntry(1).Text = .Fields("Stock")
        txtEntry(2).Text = .Fields("Short1")
        txtEntry(3).Text = .Fields("Short2")
        txtEntry(4).Text = .Fields("ICode")
        txtEntry(5).Text = .Fields("SalesPrice")
        txtEntry(6).Text = .Fields("Cost")
        txtEntry(7).Text = .Fields("ReorderPoint")
        txtEntry(8).Text = .Fields("OnHand")
        dcCategory.BoundText = .Fields![CategoryID]
    End With
    
    Exit Sub
err:
        If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    txtEntry(0).SetFocus
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewer.CUS_PK = PK
    'frmInvoiceViewer.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.lblTitle.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.show vbModal
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    If is_empty(txtEntry(1), True) = True Then Exit Sub
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("addedbyfk") = CurrUser.USER_PK
    Else
        rs.Fields("datemodified") = Now
        rs.Fields("lastuserfk") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("Barcode") = txtEntry(0).Text
        .Fields("Stock") = txtEntry(1).Text
        .Fields("Short1") = txtEntry(2).Text
        .Fields("Short2") = txtEntry(3).Text
        .Fields("ICode") = txtEntry(4).Text
        .Fields("SalesPrice") = txtEntry(5).Text
        .Fields("Cost") = txtEntry(6).Text
        .Fields("ReorderPoint") = txtEntry(7).Text
        .Fields("OnHand") = txtEntry(8).Text
        .Fields("CategoryID") = dcCategory.BoundText

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
err:
        If err.Number = -2147217887 Then Resume Next
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
    rs.Open "SELECT * FROM Stocks WHERE StockID = " & PK, CN, adOpenStatic, adLockOptimistic
    
    bind_dc "SELECT * FROM Stocks_Category", "Category", dcCategory, "CategoryID", True
    
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
            frmStocks.RefreshRecords
        ElseIf State = adStatePopupMode Then
            srcText.Text = txtEntry(0).Text
            srcText.Tag = PK
            On Error Resume Next
            srcTextAdd.Text = rs![DisplayAddr]
            srcTextCP.Text = txtEntry(6).Text
            'srcTextDisc.Text = toNumber(cmdDisc.Text)
        End If
    End If
    
    Set frmStocksAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = False
    HLText txtEntry(Index)
End Sub

Private Sub txtEntry_LostFocus(Index As Integer)
    If Index = 8 Then cmdSave.Default = True
End Sub




