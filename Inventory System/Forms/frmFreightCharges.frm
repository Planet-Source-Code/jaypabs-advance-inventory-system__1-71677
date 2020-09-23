VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFreightCharges 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Freight Charges"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5430
      TabIndex        =   18
      Top             =   4290
      Width           =   1035
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   4290
      TabIndex        =   17
      Top             =   4290
      Width           =   1035
   End
   Begin VB.ComboBox cboCreditOption 
      Height          =   315
      ItemData        =   "frmFreightCharges.frx":0000
      Left            =   2610
      List            =   "frmFreightCharges.frx":000D
      TabIndex        =   13
      Text            =   "Upon Order"
      Top             =   870
      Width           =   2265
   End
   Begin VB.ComboBox cboFreightPeriod 
      Height          =   315
      ItemData        =   "frmFreightCharges.frx":003A
      Left            =   2610
      List            =   "frmFreightCharges.frx":0044
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   495
      Width           =   2490
   End
   Begin VB.ComboBox cboFreightAgreement 
      Height          =   315
      ItemData        =   "frmFreightCharges.frx":005B
      Left            =   2610
      List            =   "frmFreightCharges.frx":006E
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   120
      Width           =   2490
   End
   Begin VB.Frame Frame1 
      Caption         =   "Expenses"
      Height          =   1185
      Left            =   1320
      TabIndex        =   0
      Top             =   2880
      Width           =   3525
      Begin VB.TextBox txtArrastre 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   2
         Text            =   "0.00"
         Top             =   630
         Width           =   1275
      End
      Begin VB.TextBox txtFreight 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   1590
         TabIndex        =   1
         Text            =   "0.00"
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Local Arrastre:"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   690
         Width           =   1035
      End
      Begin VB.Label Label19 
         Alignment       =   1  'Right Justify
         Caption         =   "Freight:"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   300
         Width           =   1035
      End
   End
   Begin MSComCtl2.DTPicker dtpDeliveryDate 
      Height          =   315
      Left            =   2610
      TabIndex        =   5
      Top             =   1650
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20709379
      CurrentDate     =   38989
   End
   Begin MSComCtl2.DTPicker dtpReceiptDate 
      Height          =   315
      Left            =   2610
      TabIndex        =   6
      Top             =   2010
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20709379
      CurrentDate     =   38989
   End
   Begin MSComCtl2.DTPicker dtpOrderDate 
      Height          =   315
      Left            =   2610
      TabIndex        =   15
      Top             =   1260
      Width           =   1785
      _ExtentX        =   3149
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   20709379
      CurrentDate     =   38989
   End
   Begin MSDataListLib.DataCombo dcShippingCompany 
      Height          =   315
      Left            =   2610
      TabIndex        =   20
      Top             =   2400
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   556
      _Version        =   393216
      Style           =   2
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Text            =   ""
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Shipping Company"
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
      Index           =   0
      Left            =   990
      TabIndex        =   19
      Top             =   2430
      Width           =   1560
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Order Date"
      Height          =   225
      Left            =   1500
      TabIndex        =   16
      Top             =   1290
      Width           =   1065
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Credit Option"
      Height          =   255
      Left            =   1380
      TabIndex        =   14
      Top             =   900
      Width           =   1185
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Period"
      Height          =   240
      Index           =   1
      Left            =   450
      TabIndex        =   12
      Top             =   495
      Width           =   2115
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight Payment Agreement"
      Height          =   240
      Index           =   6
      Left            =   450
      TabIndex        =   11
      Top             =   135
      Width           =   2115
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Delivery Date"
      Height          =   225
      Left            =   1500
      TabIndex        =   8
      Top             =   1680
      Width           =   1065
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "Receipt Date"
      Height          =   225
      Left            =   1500
      TabIndex        =   7
      Top             =   2040
      Width           =   1065
   End
End
Attribute VB_Name = "frmFreightCharges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public POID                 As Long
Public VendorPK             As Long
Public State                As FormState 'Variable used to determine on how the form used

Dim RS                      As New Recordset

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    'Connection for Vendors_Ledger
    Dim RSLedger As New Recordset

    CN.BeginTrans
    
    With RS
        If State = adStateAddMode Then
            .AddNew
            
            .Fields("POID") = POID
        End If
        
        .Fields("DeliveryDate") = dtpDeliveryDate.Value
        .Fields("ReceiptDate") = dtpReceiptDate.Value
        .Fields("ShippingCompanyID") = dcShippingCompany.BoundText
        .Fields("Freight") = txtFreight.Text
        .Fields("Arrastre") = txtArrastre.Text
        
        .Update
    End With
    
    With RSLedger
        .CursorLocation = adUseClient

        'Save freight
        .Open "SELECT * FROM Vendors_Ledger WHERE POID=" & POID & " AND BillType = 'Freight'", CN, adOpenStatic, adLockOptimistic
        
        'Save bill to Vendors_Ledger table
        If cboFreightAgreement.Text = "By supplier until freight" Then
            If cboFreightPeriod.Text = "Postpaid" Then
                .AddNew
                
                !POID = POID
            
                !VendorID = VendorPK
                !Date = Date
                !BillType = "Freight"
            
                !Credit = txtFreight.Text
                
                .Update
                
                CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, POID, [Date], Debit ) " _
                        & "VALUES (" & dcShippingCompany.BoundText & ", " & POID & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
            End If
            
            CN.Execute "INSERT INTO Local_Forwarder_Ledger ( POID, [Date], Debit ) " _
                    & "VALUES (" & POID & ",#" & Date & "#, " & txtArrastre.Text & ")"
        ElseIf cboFreightAgreement.Text = "By supplier until local arrastre" Then
            If cboFreightPeriod.Text = "Postpaid" Then
                .AddNew
                
                !POID = POID
                
                !VendorID = VendorPK
                !Date = Date
                !BillType = "Freight"
                
                !Credit = toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)
                
                .Update
                
                CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, POID, [Date], Debit ) " _
                        & "VALUES (" & dcShippingCompany.BoundText & ", " & POID & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
            End If
        ElseIf cboFreightAgreement.Text = "Half until freight" Then
            .AddNew
            
            !POID = POID
            
            !VendorID = VendorPK
            !Date = Date
            !BillType = "Freight"
            
            If cboFreightPeriod.Text = "Prepaid" Then
                !Dedit = toMoney(txtFreight.Text) / 2
            Else 'Postpaid
                !Credit = toMoney(txtFreight.Text) / 2
            
                CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, POID, [Date], Debit ) " _
                        & "VALUES (" & dcShippingCompany.BoundText & ", " & POID & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
            End If
        
            CN.Execute "INSERT INTO Local_Forwarder_Ledger ( POID, [Date], Debit ) " _
                    & "VALUES (" & POID & ",#" & Date & "#, " & txtArrastre.Text & ")"
            
            .Update
        ElseIf cboFreightAgreement.Text = "Half until local arrastre" Then
            .AddNew
            
            !POID = POID
            
            !VendorID = VendorPK
            !Date = Date
            !BillType = "Freight"
            
            If cboFreightPeriod.Text = "Prepaid" Then
                !Dedit = (toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)) / 2
            Else 'Postpaid
                !Credit = (toMoney(txtFreight.Text) + toMoney(txtArrastre.Text)) / 2
            
                CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, POID, [Date], Debit ) " _
                        & "VALUES (" & dcShippingCompany.BoundText & ", " & POID & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
            
                CN.Execute "INSERT INTO Local_Forwarder_Ledger ( POID, [Date], Debit ) " _
                        & "VALUES (" & POID & ",#" & Date & "#, " & toNumber(txtArrastre.Text) & ")"
            End If
            
            .Update
        ElseIf cboFreightAgreement.Text = "By VTM" Then
            CN.Execute "INSERT INTO Shipping_Company_Ledger ( ShippingCompanyID, POID, [Date], Debit ) " _
                    & "VALUES (" & dcShippingCompany.BoundText & ", " & POID & ", #" & Date & "#, " & toNumber(txtFreight.Text) & ")"
        
            CN.Execute "INSERT INTO Local_Forwarder_Ledger ( POID, [Date], Debit ) " _
                    & "VALUES (" & POID & ",#" & Date & "#, " & toNumber(txtArrastre.Text) & ")"
        End If
    End With
    
    
    CN.CommitTrans
    
    Set RS = Nothing
    Set RSLedger = Nothing
    
End Sub

Private Sub Form_Load()
    Dim RSPO As New Recordset

    With RSPO
        .Open "SELECT * FROM qry_Purchase_Order WHERE POID=" & POID, CN, adOpenStatic, adLockOptimistic
        
        cboFreightAgreement.Text = .Fields("FreightAgreement")
        cboFreightPeriod.Text = .Fields("FreightPeriod")
        cboCreditOption.Text = .Fields("CreditOption")
        dtpOrderDate.Value = .Fields("Date")
        
        .Close
    End With
    
    bind_dc "SELECT * FROM Shipping_Company", "ShippingCompany", dcShippingCompany, "ShippingCompanyID", True

    With RS
        .Open "SELECT * FROM Freight_Charges WHERE POID=" & POID, CN, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            State = adStateEditMode
            
            dtpDeliveryDate.Value = .Fields("DeliveryDate")
            dtpReceiptDate.Value = .Fields("ReceiptDate")
            dcShippingCompany.BoundText = .Fields("ShippingCompanyID")
            txtFreight.Text = .Fields("Freight")
            txtArrastre.Text = .Fields("Arrastre")
        Else
            dtpDeliveryDate.Value = Date
            dtpReceiptDate.Value = Date
        
            State = adStateAddMode
        End If
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmFreightCharges = Nothing
End Sub
