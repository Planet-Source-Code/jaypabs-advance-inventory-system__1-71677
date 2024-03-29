VERSION 5.00
Begin VB.Form frmBusinessInfo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Business Information"
   ClientHeight    =   2670
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmBusinessInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2670
   ScaleWidth      =   6090
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1590
      TabIndex        =   0
      Top             =   240
      Width           =   4335
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   -450
      TabIndex        =   7
      Top             =   1920
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   53
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1590
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   630
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1590
      TabIndex        =   2
      Top             =   1320
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Changes"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2910
      TabIndex        =   3
      Top             =   2160
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   4710
      TabIndex        =   4
      Top             =   2160
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Business Name:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Business Address:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   6
      Top             =   630
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Contact Info:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   1320
      Width           =   1620
   End
End
Attribute VB_Name = "frmBusinessInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rs_set As New Recordset

Private Sub Command1_Click()
    If is_empty(Text1) = True Then Exit Sub
    If is_empty(Text2) = True Then Exit Sub
    
    With rs_set
        .Fields("Name") = Text1.Text
        .Fields("Address") = Text2.Text
        .Fields("ContactInfo") = Text3.Text
        .Update
    End With
    
    With CurrBiz
        .BUSINESS_NAME = Text1.Text
        .BUSINESS_ADDRESS = Text2.Text
        .BUSINESS_CONTACT_INFO = Text3.Text
    End With
    MsgBox "Changes has been successfully saved.", vbInformation
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
    rs_set.Open "SELECT * FROM TBL_BUSINESS_INFO", CN, adOpenStatic, adLockOptimistic
    
    Text1.Text = rs_set.Fields("Name")
    Text2.Text = rs_set.Fields("Address")
    Text3.Text = rs_set.Fields("ContactInfo")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmBusinessInfo = Nothing
End Sub

Private Sub Text1_GotFocus()
    HLText Text1
End Sub

Private Sub Text2_GotFocus()
    HLText Text2
End Sub

