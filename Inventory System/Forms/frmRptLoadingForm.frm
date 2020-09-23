VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmRptLoadingForm 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   LinkTopic       =   "Form1"
   ScaleHeight     =   2235
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print"
      Height          =   345
      Left            =   1800
      TabIndex        =   1
      Top             =   1680
      Width           =   1065
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   345
      Left            =   3000
      TabIndex        =   0
      Top             =   1680
      Width           =   915
   End
   Begin MSDataListLib.DataCombo dcRoute 
      Height          =   315
      Left            =   1770
      TabIndex        =   2
      Top             =   480
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   1770
      TabIndex        =   3
      Top             =   840
      Width           =   2205
      _ExtentX        =   3889
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89587715
      CurrentDate     =   38207
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Route"
      Height          =   225
      Left            =   450
      TabIndex        =   5
      Top             =   510
      Width           =   1275
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date:"
      Height          =   255
      Index           =   1
      Left            =   420
      TabIndex        =   4
      Top             =   870
      Width           =   1305
   End
   Begin VB.Shape Shape1 
      Height          =   2115
      Left            =   60
      Top             =   60
      Width           =   4095
   End
End
Attribute VB_Name = "frmRptLoadingForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    Me.Hide
    
    Unload frmReports

    With frmReports
        .strReport = "Loading Form Report"
        .strWhere = "{qry_rpt_Loading_Form.Route} = '" & dcRoute.BoundText & "' " _
                            & "AND {qry_rpt_Loading_Form.DeliveryDate} = #" & dtpDate.Value & "#"
        
        LoadForm frmReports
    End With
    
    Unload Me
End Sub

Private Sub Form_Load()
    bind_dc "SELECT Route, Desc FROM Routes", "Desc", dcRoute, "Route", True
    
    dtpDate.Value = Date
End Sub


