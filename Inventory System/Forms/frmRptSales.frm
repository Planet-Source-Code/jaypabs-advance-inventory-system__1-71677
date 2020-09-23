VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptSales 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sales Report"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   4425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   465
      Left            =   1110
      TabIndex        =   1
      Top             =   1230
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   2550
      TabIndex        =   0
      Top             =   1230
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtpBegDate 
      Height          =   375
      Left            =   2250
      TabIndex        =   2
      Top             =   150
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89587715
      CurrentDate     =   39156
   End
   Begin MSComCtl2.DTPicker dtpEndDate 
      Height          =   375
      Left            =   2250
      TabIndex        =   3
      Top             =   600
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89587715
      CurrentDate     =   39156
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Beginning Date"
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   150
      Width           =   1395
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date"
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   600
      Width           =   1395
   End
End
Attribute VB_Name = "frmRptSales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    dtpBegDate.Value = Date
    dtpEndDate.Value = Date
End Sub

Private Sub cmdPreview_Click()
    Me.Hide
    
    Unload frmReports

    With frmReports
        .strReport = "Sales Report"
        .strWhere = "{qry_rpt_Sales.DeliveryDate} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        
        LoadForm frmReports
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub


