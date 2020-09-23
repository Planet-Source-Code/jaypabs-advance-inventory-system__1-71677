VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTransferQty 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer Qty"
   ClientHeight    =   3585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3585
   ScaleWidth      =   5910
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   4830
      TabIndex        =   14
      Top             =   3120
      Width           =   765
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transfer QTY"
      Height          =   2175
      Left            =   2610
      TabIndex        =   10
      Top             =   780
      Width           =   3015
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1290
         TabIndex        =   23
         Text            =   "0"
         Top             =   300
         Width           =   915
      End
      Begin VB.CommandButton cmdTransfer 
         Caption         =   "&Transfer"
         Height          =   285
         Left            =   1890
         TabIndex        =   13
         Top             =   1740
         Width           =   885
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Index           =   0
         Left            =   1290
         TabIndex        =   17
         Top             =   720
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Index           =   1
         Left            =   1290
         TabIndex        =   18
         Top             =   1170
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Qty:"
         Height          =   285
         Left            =   510
         TabIndex        =   24
         Top             =   300
         Width           =   585
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "To:"
         Height          =   285
         Left            =   300
         TabIndex        =   12
         Top             =   1200
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "From:"
         Height          =   285
         Left            =   300
         TabIndex        =   11
         Top             =   720
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Current QTY"
      Height          =   2205
      Left            =   270
      TabIndex        =   1
      Top             =   750
      Width           =   2175
      Begin VB.TextBox txtUnit 
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   960
         TabIndex        =   5
         Top             =   1710
         Width           =   705
      End
      Begin VB.TextBox txtUnit 
         Enabled         =   0   'False
         Height          =   315
         Index           =   3
         Left            =   960
         TabIndex        =   4
         Top             =   1350
         Width           =   705
      End
      Begin VB.TextBox txtUnit 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   960
         TabIndex        =   3
         Top             =   990
         Width           =   705
      End
      Begin VB.TextBox txtUnit 
         Height          =   315
         Index           =   1
         Left            =   960
         TabIndex        =   2
         Top             =   630
         Width           =   705
      End
      Begin VB.Label Label6 
         Caption         =   "4"
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   22
         Top             =   1710
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "3"
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   21
         Top             =   1350
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "2"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   20
         Top             =   990
         Width           =   195
      End
      Begin VB.Label Label6 
         Caption         =   "1"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   19
         Top             =   630
         Width           =   195
      End
      Begin VB.Label Label5 
         Caption         =   "Onhand QTY"
         Height          =   255
         Left            =   1050
         TabIndex        =   16
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label4 
         Caption         =   "Packaging"
         Height          =   255
         Left            =   150
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "N/A"
         Height          =   225
         Index           =   4
         Left            =   300
         TabIndex        =   9
         Top             =   1740
         Width           =   585
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "N/A"
         Height          =   225
         Index           =   3
         Left            =   300
         TabIndex        =   8
         Top             =   1380
         Width           =   585
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "N/A"
         Height          =   225
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   990
         Width           =   585
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Caption         =   "Unit"
         Height          =   225
         Index           =   1
         Left            =   300
         TabIndex        =   6
         Top             =   630
         Width           =   585
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Insufficient qty. Please adjust qty below."
      BeginProperty Font 
         Name            =   "Lucida Sans Unicode"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   285
      TabIndex        =   0
      Top             =   180
      Width           =   5355
   End
End
Attribute VB_Name = "frmTransferQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public StockID      As Integer

Dim rs As New Recordset

Private Sub cmdClose_Click()
    rs.Close
    Set rs = Nothing
    
    Unload Me
End Sub

Private Sub cmdTransfer_Click()
    Dim intQtyToAdd As Integer
    
    If toNumber(txtQty.Text) < 1 Then
        txtQty.SetFocus
        
        Exit Sub
    End If
    
    If MsgBox("This will transfer qty from " & dcUnit(0).Text & " to " & dcUnit(1).Text & vbCrLf & vbCrLf & "Do you want to continue?", vbInformation + vbYesNo) = vbNo Then Exit Sub
    
    If dcUnit(0).Text = dcUnit(1).Text Then
        MsgBox "Please select a different packaging in From and To field", vbExclamation
        
        Exit Sub
    End If
    
    If dcUnit(0).BoundText > dcUnit(1).BoundText Then
        MsgBox "You can not subtract from the lowest packaging unit.", vbExclamation
        
        Exit Sub
    End If
    
    If dcUnit(0).BoundText < dcUnit(1).BoundText Then
        rs.MoveFirst
        
        rs.Find "[Order] = " & dcUnit(0).BoundText
        
        If rs![Onhand] <= 0 Then
            MsgBox "There is no qty left from " & dcUnit(0).Text & " unit.", vbExclamation
            
            Exit Sub
        ElseIf toNumber(txtQty.Text) > rs![Onhand] Then
            MsgBox "Qty given is greater than the qty left.", vbExclamation
            
            Exit Sub
        Else
            On Error GoTo err_cmdTransfer_Click
            
            CN.BeginTrans
            
            'Deduct from the highest packaging unit
            rs!Onhand = rs!Onhand - toNumber(txtQty.Text)
            rs.Update
            
            intQtyToAdd = SubtractFromHighestQty(dcUnit(0).BoundText, dcUnit(1).BoundText, txtQty.Text, rs)
            'Add from the highest packaging unit
            rs!Onhand = rs!Onhand + intQtyToAdd
            rs.Update
            
            CN.CommitTrans
            
            MsgBox "Qty sucessfully transferred.", vbInformation
            
            txtUnit(dcUnit(0).BoundText).Text = txtUnit(dcUnit(0).BoundText) - toNumber(txtQty.Text)
            txtUnit(dcUnit(1).BoundText) = txtUnit(rs!Order) + intQtyToAdd
        End If
    End If
    
    Exit Sub
    
err_cmdTransfer_Click:
    CN.RollbackTrans
    MsgBox err.Number & " " & err.Description
End Sub

'This function is used to subtract qty from the highest packaging unit
Private Function SubtractFromHighestQty(Order As Integer, ByVal Ordertmp As Integer, intQty As Integer, rs As Recordset)
    'Order variable is the highest packaging unit
    'Ordertmp variable is the lowest packaging unit
    SubtractFromHighestQty = intQty
    Do Until Order = Ordertmp
        Order = Order + 1
        
        rs.MoveNext
        
        SubtractFromHighestQty = SubtractFromHighestQty * rs!Qty
    Loop
End Function

'Private Sub dcUnit_Click(Index As Integer, Area As Integer)
'    If dcUnit(0).BoundText <> dcUnit(1).BoundText Then
'        If dcUnit(0).BoundText > dcUnit(1).BoundText Then
'            rs.MoveFirst
'
'            rs.Find "[Order] = " & dcUnit(1).BoundText
'
'            Debug.Print "suggested qty: " & SubtractFromHighestQty(dcUnit(1).BoundText, dcUnit(0).BoundText, 1, rs)
'        End If
'    End If
'End Sub

Private Sub Form_Load()
    Dim I As Integer
    
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & StockID, "Unit", dcUnit(0), "Order", True
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & StockID, "Unit", dcUnit(1), "Order", True
    
    rs.Open "SELECT StockUnitID, [Order], Qty, UnitID, Unit, Onhand FROM qry_Stock_Unit WHERE StockID=" & StockID & " ORDER BY [Order] ASC", CN, adOpenStatic, adLockOptimistic
        
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        For I = 1 To rs.RecordCount
            lblUnit(I).Caption = rs!Unit
            txtUnit(I).Text = rs!Onhand
            txtUnit(I).Enabled = True
            
            rs.MoveNext
        Next I
    End If
End Sub

Private Sub txtQty_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub
