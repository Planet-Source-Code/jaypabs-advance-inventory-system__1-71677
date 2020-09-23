VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSalesReceiptsBatchAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Receipt by Batch"
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   6075
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   0
      Left            =   1770
      TabIndex        =   1
      Top             =   575
      Width           =   1155
   End
   Begin MSDataListLib.DataCombo dcBooking 
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   910
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo dcRoute 
      Height          =   315
      Left            =   1770
      TabIndex        =   0
      Top             =   210
      Width           =   3345
      _ExtentX        =   5900
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   2
      Left            =   1770
      TabIndex        =   6
      Top             =   2310
      Width           =   2475
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   315
      Left            =   2850
      TabIndex        =   8
      Top             =   2850
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   4305
      TabIndex        =   9
      Top             =   2850
      Width           =   1335
   End
   Begin VB.CommandButton cmdUsrHistory 
      Caption         =   "Modification History"
      Height          =   315
      Left            =   420
      TabIndex        =   7
      Top             =   2880
      Width           =   1680
   End
   Begin VB.TextBox txtEntry 
      Height          =   285
      Index           =   1
      Left            =   1770
      TabIndex        =   5
      Top             =   1975
      Width           =   2475
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   150
      TabIndex        =   10
      Top             =   2730
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   53
   End
   Begin MSDataListLib.DataCombo dcCollection 
      Height          =   315
      Left            =   1770
      TabIndex        =   3
      Top             =   1275
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   556
      _Version        =   393216
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   285
      Left            =   1770
      TabIndex        =   4
      Top             =   1640
      Width           =   2475
      _ExtentX        =   4366
      _ExtentY        =   503
      _Version        =   393216
      Format          =   44695553
      CurrentDate     =   39166
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Truck No"
      Height          =   255
      Left            =   600
      TabIndex        =   17
      Top             =   560
      Width           =   1095
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Helper"
      Height          =   255
      Left            =   600
      TabIndex        =   16
      Top             =   2310
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   255
      Left            =   600
      TabIndex        =   15
      Top             =   210
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Booking"
      Height          =   255
      Left            =   600
      TabIndex        =   14
      Top             =   910
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Collection"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date"
      Height          =   255
      Left            =   600
      TabIndex        =   12
      Top             =   1610
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Driver"
      Height          =   255
      Left            =   600
      TabIndex        =   11
      Top             =   1960
      Width           =   1095
   End
End
Attribute VB_Name = "frmSalesReceiptsBatchAE"
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
        dcRoute.BoundText = .Fields("RouteID")
        txtEntry(0).Text = .Fields("TruckNo")
        dcBooking.BoundText = .Fields("BookingAgent")
        dcCollection.BoundText = .Fields("CollectionAgent")
        dtpDate.Value = .Fields("DeliveryDate")
        txtEntry(1).Text = .Fields("Driver")
        txtEntry(2).Text = .Fields("Helper")
    End With
    dcRoute.SetFocus
    
    Exit Sub
    
errHandler:
    If err.Number = 94 Then Resume Next
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub ResetFields()
    clearText Me
    
    dcRoute.Text = ""
    txtEntry(0).Text = ""
    dcBooking.Text = ""
    dcCollection.Text = ""
    txtEntry(1).Text = ""
    txtEntry(2).Text = ""

    dcRoute.SetFocus
End Sub

Private Sub cmdPH_Click()
    'frmInvoiceViewer.CUS_PK = PK
    'frmInvoiceViewer.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.lblTitle.Caption = "Purchase History Viewer"
    'frmInvoiceViewer.show vbModal
End Sub

Private Sub cmdSave_Click()
On Error GoTo err

    If dcRoute.Text = "" Then Exit Sub
        
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.AddNew
        rs.Fields("AddedByFK") = CurrUser.USER_PK
    Else
        rs.Fields("DateModified") = Now
        rs.Fields("LastUserFK") = CurrUser.USER_PK
    End If
    
    With rs
        .Fields("RouteID") = dcRoute.BoundText
        .Fields("TruckNo") = txtEntry(0).Text
        .Fields("BookingAgent") = dcBooking.BoundText
        .Fields("CollectionAgent") = dcCollection.BoundText
        .Fields("DeliveryDate") = dtpDate.Value
        .Fields("Driver") = txtEntry(1).Text
        .Fields("Helper") = txtEntry(2).Text
        
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
  & "Form: frmReceiptBatchAE" & vbCr _
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
   
    bind_dc "SELECT * FROM Routes", "Desc", dcRoute, "RouteID", True
    bind_dc "SELECT * FROM Agents", "AgentName", dcBooking, "AgentID", True
    bind_dc "SELECT * FROM Agents", "AgentName", dcCollection, "AgentID", True
          
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM Receipts_Batch WHERE ReceiptBatchID = " & PK, CN, adOpenStatic, adLockOptimistic
        
        Caption = "Create New Entry"
        cmdUsrHistory.Enabled = False
        
        dtpDate.Value = Date
    Else
        rs.CursorLocation = adUseClient
        rs.Open "SELECT * FROM Receipts_Batch WHERE ReceiptBatchID = " & PK, CN, adOpenStatic, adLockOptimistic
    
        Caption = "Edit Entry"
        DisplayForEditing
        'cmdPH.Enabled = True
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        If State = adStateAddMode Or State = adStateEditMode Then
            frmSalesReceiptsBatch.RefreshRecords1
        End If
    End If
    
    Set frmSalesReceiptsBatchAE = Nothing
End Sub

Private Sub txtEntry_GotFocus(Index As Integer)
    HLText txtEntry(Index)
End Sub

