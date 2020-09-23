VERSION 5.00
Begin VB.Form frmWeeklyInv 
   BorderStyle     =   0  'None
   Caption         =   "Weekly Inventory"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6840
   LinkTopic       =   "Form2"
   ScaleHeight     =   4290
   ScaleWidth      =   6840
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5700
      TabIndex        =   0
      Top             =   3720
      Width           =   855
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   6
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   5
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2670
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   4
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2220
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   3
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1770
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   2
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1335
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.TextBox txtQty 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   375
      Index           =   1
      Left            =   3450
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   885
      Visible         =   0   'False
      Width           =   1875
   End
   Begin Inventory.ctrlLiner ctrlLiner1 
      Height          =   30
      Left            =   270
      TabIndex        =   15
      Top             =   3540
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   53
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 1"
      Height          =   255
      Index           =   11
      Left            =   840
      TabIndex        =   21
      Top             =   900
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 2"
      Height          =   255
      Index           =   10
      Left            =   840
      TabIndex        =   20
      Top             =   1320
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 3"
      Height          =   255
      Index           =   9
      Left            =   840
      TabIndex        =   19
      Top             =   1770
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 4"
      Height          =   255
      Index           =   8
      Left            =   840
      TabIndex        =   18
      Top             =   2220
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 5"
      Height          =   255
      Index           =   7
      Left            =   840
      TabIndex        =   17
      Top             =   2670
      Width           =   1065
   End
   Begin VB.Label lblLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Week 6"
      Height          =   255
      Index           =   0
      Left            =   840
      TabIndex        =   16
      Top             =   3120
      Width           =   1065
   End
   Begin VB.Shape Shape1 
      Height          =   4155
      Left            =   90
      Top             =   60
      Width           =   6615
   End
   Begin VB.Label lblStatus 
      Caption         =   "Status"
      Height          =   375
      Left            =   660
      TabIndex        =   14
      Top             =   3690
      Width           =   4215
   End
   Begin VB.Label lblProduct 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   540
      TabIndex        =   13
      Top             =   210
      Width           =   5655
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   6
      Left            =   2010
      TabIndex        =   11
      Top             =   3120
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   5
      Left            =   2010
      TabIndex        =   9
      Top             =   2670
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   4
      Left            =   2010
      TabIndex        =   7
      Top             =   2220
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   3
      Left            =   2010
      TabIndex        =   5
      Top             =   1770
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   2
      Left            =   2010
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Label1"
      Height          =   255
      Index           =   1
      Left            =   2010
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H80000000&
      BackStyle       =   1  'Opaque
      Height          =   615
      Left            =   300
      Top             =   150
      Width           =   6255
   End
End
Attribute VB_Name = "frmWeeklyInv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public intProductID As Integer

Private Sub CmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
On Error GoTo erR

    Dim RSWeeklyInv As New Recordset
    Dim I As Integer
    Dim dDate As Date
    
    lblProduct.Caption = getValueAt("SELECT StockID,Stock FROM Stocks WHERE StockID=" & intProductID, "Stock")
    
    dDate = Date - 42
    
    Do Until Format(dDate, "dddd") = "Monday"
        dDate = dDate - 1
    Loop

    With RSWeeklyInv
        .CursorLocation = adUseClient

        .Open "TRANSFORM Count(qry_Weekly_Inventory.StockCardID) AS CountOfStockCardID " _
            & "SELECT DatePart('ww',[DateInsert],1) AS Offtake, " _
            & "Count(qry_Weekly_Inventory.StockCardID) AS [Total Of StockCardID] " _
            & "From qry_Weekly_Inventory " _
            & "Where (((qry_Weekly_Inventory.StockID) = " & intProductID & ")) AND " _
            & "DateInsert >= #" & dDate & "# " _
            & "GROUP BY DatePart('ww',[DateInsert],1) " _
            & "PIVOT qry_Weekly_Inventory.StockID", CN, adOpenStatic, adLockOptimistic
 
        I = 1
        If .RecordCount > 0 Then
            Do While Not .EOF
                lblLabel(I).Caption = DateAdd("ww", ![Offtake], "1/1/" & Year(Date)) - 7
                lblLabel(I).Caption = Format(lblLabel(I).Caption, "d-mmm-yyyy")
                txtQty(I).Text = .Fields(2) 'Fields(2) = StockID
                
                lblLabel(I).Visible = True
                txtQty(I).Visible = True
                
                I = I + 1
                .MoveNext
            Loop
            lblStatus.Caption = "There are " & .RecordCount & " week(s) offtake for this product"
        Else
            lblStatus.Caption = "No offtake for this product"
        End If
        
        .Close
        Set RSWeeklyInv = Nothing
    End With
    
    Exit Sub
    
erR:
    MsgBox erR.Number & " " & erR.Description, vbCritical
End Sub

