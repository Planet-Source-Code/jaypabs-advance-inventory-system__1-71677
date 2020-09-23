VERSION 5.00
Begin VB.Form frmSuggestedQty 
   BorderStyle     =   0  'None
   Caption         =   "Form2"
   ClientHeight    =   3510
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9405
   LinkTopic       =   "Form2"
   ScaleHeight     =   3510
   ScaleWidth      =   9405
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDeduct 
      Caption         =   "Deduct"
      Height          =   435
      Left            =   5190
      TabIndex        =   11
      Top             =   2820
      Width           =   1275
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   435
      Left            =   7890
      TabIndex        =   9
      Top             =   2820
      Width           =   1275
   End
   Begin VB.CheckBox Check2 
      Enabled         =   0   'False
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   1470
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.TextBox txtProduct 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   1
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Width           =   7575
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   1
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CheckBox chkSuggested 
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   2220
      Width           =   195
   End
   Begin VB.TextBox txtProduct 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1410
      Width           =   7575
   End
   Begin VB.TextBox txtQty 
      Appearance      =   0  'Flat
      Height          =   345
      Index           =   0
      Left            =   540
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1410
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "O&K"
      Height          =   435
      Left            =   6540
      TabIndex        =   0
      Top             =   2820
      Width           =   1275
   End
   Begin VB.Label Label2 
      Caption         =   "Suggested Qty:"
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1800
      Width           =   1245
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "To deduct from previous Order click here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   375
      Left            =   720
      TabIndex        =   10
      Top             =   2850
      Width           =   4395
   End
   Begin VB.Shape Shape2 
      Height          =   3315
      Left            =   90
      Top             =   60
      Width           =   9255
   End
   Begin VB.Label lblHeader2 
      BackStyle       =   0  'Transparent
      Caption         =   "Below is the list of suggested Qty for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   270
      TabIndex        =   2
      Top             =   750
      Width           =   8805
   End
   Begin VB.Label lblHeader1 
      BackStyle       =   0  'Transparent
      Caption         =   "Insufficient Qty for:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   300
      TabIndex        =   1
      Top             =   270
      Width           =   8775
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      Height          =   1065
      Left            =   180
      Top             =   180
      Width           =   9075
   End
End
Attribute VB_Name = "frmSuggestedQty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intStockID           As Integer
Public strProduct           As String
Public intQtyOrdered        As Integer
Public intQtySuggested      As Integer
Public blnUseSuggestedQty   As Boolean
Public blnCancel            As Boolean

Private Sub cmdCancel_Click()
    blnCancel = True
    Unload Me
End Sub

Private Sub cmdDeduct_Click()
    With frmCustomersItem
        .StockID = intStockID
        
        .show 1
    End With
End Sub

Private Sub cmdOK_Click()
    If chkSuggested.Value = 1 Then
        blnUseSuggestedQty = True
    Else
        blnUseSuggestedQty = False
    End If
    
    If intQtyOrdered = 0 And chkSuggested.Value = 0 Then
        blnCancel = True
    Else
        blnCancel = False
    End If
    
    Unload Me
    
End Sub

Private Sub Form_Load()
    lblHeader1.Caption = lblHeader1.Caption & " " & strProduct
    lblHeader2.Caption = lblHeader2.Caption & " " & strProduct
    
    txtQty(0).Text = intQtyOrdered
    txtQty(1).Text = intQtySuggested
    
    txtProduct(0).Text = strProduct
    txtProduct(1).Text = strProduct
End Sub
