VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmRptCustomers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customers Report"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5790
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   5790
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPreview 
      Caption         =   "&Preview"
      Height          =   465
      Left            =   2790
      TabIndex        =   1
      Top             =   1740
      Width           =   1305
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   465
      Left            =   4230
      TabIndex        =   0
      Top             =   1740
      Width           =   1305
   End
   Begin MSComCtl2.DTPicker dtpBegDate 
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   750
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
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   1605
      _ExtentX        =   2831
      _ExtentY        =   661
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   89587715
      CurrentDate     =   39156
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdClient 
      Height          =   315
      Left            =   1860
      TabIndex        =   7
      Top             =   270
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   556
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Beginning Date"
      Height          =   315
      Left            =   450
      TabIndex        =   6
      Top             =   750
      Width           =   1935
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "End Date"
      Height          =   315
      Left            =   450
      TabIndex        =   5
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Customer"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   4
      Top             =   270
      Width           =   1545
   End
End
Attribute VB_Name = "frmRptCustomers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    InitNSD
    
    dtpBegDate.Value = Date
    dtpEndDate.Value = Date
End Sub

Private Sub cmdPreview_Click()
    Me.Hide
        
    Unload frmReports

    With frmReports
        .strReport = "Customers Report"

        If nsdClient.Text = "" Then
            .strWhere = "{qry_rpt_Customers.DateIssued} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        Else
            .strWhere = "{qry_rpt_Customers.ClientID} = " & nsdClient.BoundText & " " _
                                & "AND {qry_rpt_Customers.DateIssued} IN #" & dtpBegDate.Value & "# TO #" & dtpEndDate.Value & "#"
        End If

        LoadForm frmReports
    End With
    
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub InitNSD()
    'For Client
    With nsdClient
        .ClearColumn
        .AddColumn "Client ID", 800
        .AddColumn "Company", 2264.88
        .AddColumn "City", 2670.23
        .AddColumn "Owner's Name", 2670.23
        .AddColumn "Credit Term", 0
        .Connection = CN.ConnectionString
        
        .sqlFields = "ClientID, Company, City, OwnersName, CreditTerm"
        .sqlTables = "qry_Clients1"
        
        .sqlSortOrder = "Company ASC"
        
        .BoundField = "ClientID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 7000, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Clients Record"
        
    End With
End Sub
