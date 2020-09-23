VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{FF49E21B-EA30-11D9-85DF-812F544F012A}#69.0#0"; "ctrlNSDataCombo.ocx"
Begin VB.Form frmAssortedProductAE 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Assorted Product"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   9420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdTasks 
      Caption         =   "Assorted Product Tasks"
      Height          =   315
      Left            =   5550
      TabIndex        =   11
      Top             =   5520
      Width           =   1995
   End
   Begin VB.CommandButton btnRemove 
      Height          =   275
      Left            =   270
      Picture         =   "frmAssortedProductAE.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Remove"
      Top             =   2970
      Visible         =   0   'False
      Width           =   275
   End
   Begin VB.TextBox txtNotes 
      Height          =   1335
      Left            =   180
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   12
      Tag             =   "Remarks"
      Top             =   5790
      Width           =   4980
   End
   Begin VB.ComboBox cboStatus 
      Height          =   315
      ItemData        =   "frmAssortedProductAE.frx":01B2
      Left            =   1290
      List            =   "frmAssortedProductAE.frx":01BC
      TabIndex        =   4
      Text            =   "On Hold"
      Top             =   1380
      Width           =   2325
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   315
      Left            =   7620
      TabIndex        =   9
      Top             =   5520
      Width           =   855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   315
      Left            =   8535
      TabIndex        =   10
      Top             =   5520
      Width           =   795
   End
   Begin VB.PictureBox picPurchase 
      BorderStyle     =   0  'None
      Height          =   630
      Left            =   240
      ScaleHeight     =   630
      ScaleWidth      =   9105
      TabIndex        =   13
      Top             =   1860
      Width           =   9105
      Begin VB.TextBox txtQty 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4125
         TabIndex        =   6
         Text            =   "0"
         Top             =   240
         Width           =   660
      End
      Begin VB.CommandButton btnAdd 
         Caption         =   "Add"
         Enabled         =   0   'False
         Height          =   315
         Left            =   6030
         TabIndex        =   8
         Top             =   270
         Width           =   840
      End
      Begin ctrlNSDataCombo.NSDataCombo nsdStock 
         Height          =   315
         Left            =   0
         TabIndex        =   5
         Top             =   225
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo dcUnit 
         Height          =   315
         Left            =   4830
         TabIndex        =   7
         Top             =   240
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Text            =   ""
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Qty"
         Height          =   240
         Index           =   10
         Left            =   4125
         TabIndex        =   16
         Top             =   0
         Width           =   660
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Items/Stocks"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000011D&
         Height          =   240
         Index           =   8
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label Labels 
         BackStyle       =   0  'Transparent
         Caption         =   "Unit"
         Height          =   240
         Index           =   2
         Left            =   4830
         TabIndex        =   14
         Top             =   0
         Width           =   900
      End
   End
   Begin VB.TextBox txtQtyParent 
      Height          =   315
      Left            =   1290
      TabIndex        =   1
      Top             =   600
      Width           =   2115
   End
   Begin ctrlNSDataCombo.NSDataCombo nsdProduct 
      Height          =   315
      Left            =   1290
      TabIndex        =   0
      Top             =   210
      Width           =   5100
      _ExtentX        =   8996
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   2610
      Left            =   210
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2820
      Width           =   9105
      _ExtentX        =   16060
      _ExtentY        =   4604
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   275
      ForeColorFixed  =   -2147483640
      BackColorSel    =   1091552
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
      GridColorUnpopulated=   -2147483633
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   1290
      TabIndex        =   3
      Top             =   990
      Width           =   2505
   End
   Begin MSComCtl2.DTPicker dtpDate 
      Height          =   315
      Left            =   1290
      TabIndex        =   2
      Top             =   990
      Width           =   2505
      _ExtentX        =   4419
      _ExtentY        =   556
      _Version        =   393216
      CustomFormat    =   "MMM-dd-yyyy"
      Format          =   44695555
      CurrentDate     =   38207
   End
   Begin VB.Label Labels 
      Alignment       =   1  'Right Justify
      Caption         =   "Date"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   24
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   210
      Left            =   270
      TabIndex        =   23
      Top             =   2520
      Width           =   4365
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000014&
      Index           =   1
      X1              =   180
      X2              =   9300
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      BorderWidth     =   2
      Index           =   1
      X1              =   180
      X2              =   9300
      Y1              =   1770
      Y2              =   1770
   End
   Begin VB.Label Labels 
      Caption         =   "Notes"
      Height          =   240
      Index           =   4
      Left            =   180
      TabIndex        =   21
      Top             =   5520
      Width           =   990
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Status"
      Height          =   255
      Left            =   210
      TabIndex        =   20
      Top             =   1410
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Qty"
      Height          =   255
      Left            =   210
      TabIndex        =   19
      Top             =   630
      Width           =   1035
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Product"
      Height          =   255
      Left            =   210
      TabIndex        =   18
      Top             =   210
      Width           =   1035
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H80000010&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000010&
      Height          =   240
      Left            =   210
      Top             =   2520
      Width           =   9105
   End
   Begin VB.Menu mnu_Tasks 
      Caption         =   "Tasks"
      Visible         =   0   'False
      Begin VB.Menu mnu_History 
         Caption         =   "History"
      End
   End
End
Attribute VB_Name = "frmAssortedProductAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public State                As FormState 'Variable used to determine on how the form used
Public PK                   As Long 'Variable used to get what record is going to edit
Public CloseMe              As Boolean
Public ForCusAcc            As Boolean

Dim cIRowCount              As Integer

Dim HaveAction              As Boolean 'Variable used to detect if the user perform some action
Dim rs                      As New Recordset 'Main recordset for Invoice
Dim blnSave                 As Boolean
Dim intQtyOld               As Integer 'Old txtQty Value. Hold when editing qty

Private Sub btnAdd_Click()
On Error GoTo erR
    
    Dim RSStockUnit As New Recordset
    
    Dim intTotalOnhand          As Integer
    Dim intTotalIncoming        As Integer
    Dim intTotalOnhInc          As Integer 'Total of Onhand + Incoming
    Dim intExcessQty            As Integer
    
    Dim intSuggestedQty         As Integer
    Dim blnAddIncoming          As Boolean
    Dim intQtyOrdered           As Integer 'hold the value of txtQty
    Dim intCount                As Integer
    
    If nsdStock.Text = "" Then nsdStock.SetFocus: Exit Sub
    
    If dcUnit.Text = "" Then
        MsgBox "Please select unit", vbInformation
        dcUnit.SetFocus
        Exit Sub
    End If
    
    Dim CurrRow As Integer

    Dim intStockID As Integer
    
    CurrRow = getFlexPos(Grid, 4, nsdStock.Tag)
    intStockID = nsdStock.Tag
    
    RSStockUnit.CursorLocation = adUseClient
    RSStockUnit.Open "SELECT * FROM qry_Stock_Unit WHERE StockID =" & intStockID & " ORDER BY Stock_Unit.Order ASC", CN, adOpenStatic, adLockOptimistic
    
    intQtyOrdered = txtQty.Text
              
    RSStockUnit.Find "UnitID = " & dcUnit.BoundText
          
    If RSStockUnit!Onhand < intQtyOrdered Then GoSub GetOnhand

Continue:
    'Save to stock card
'    Dim RSStockCard As New Recordset
'
'    RSStockCard.CursorLocation = adUseClient
'    RSStockCard.Open "SELECT * FROM Stock_Card", CN, adOpenStatic, adLockOptimistic

    'Add to grid
    With Grid
        If CurrRow < 0 Then
            'Perform if the record is not exist
            If .Rows = 2 And .TextMatrix(1, 4) = "" Then
                .TextMatrix(1, 1) = nsdStock.Text
                .TextMatrix(1, 2) = intQtyOrdered 'txtQty.Text
                .TextMatrix(1, 3) = dcUnit.Text
                .TextMatrix(1, 4) = intStockID
            Else
AddIncoming:
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 1) = nsdStock.Text
                .TextMatrix(.Rows - 1, 2) = intQtyOrdered 'txtQty.Text
                .TextMatrix(.Rows - 1, 3) = dcUnit.Text
                .TextMatrix(.Rows - 1, 4) = intStockID
                
                .FillStyle = 1

                .Row = .Rows - 1
                .ColSel = 3
                If blnAddIncoming = True And intCount = 2 Then
                    .CellForeColor = vbBlue
                    
                    blnAddIncoming = False
                End If
            End If
            'Increase the record count
            cIRowCount = cIRowCount + 1
        Else
            If MsgBox("Item already exist. Do you want to replace it?", vbQuestion + vbYesNo) = vbYes Then
                .Row = CurrRow
                                
                .TextMatrix(CurrRow, 1) = nsdStock.Text
                .TextMatrix(CurrRow, 2) = intQtyOrdered 'txtQty.Text
                .TextMatrix(CurrRow, 3) = dcUnit.Text
             
                'deduct qty from Stock Unit's table
                RSStockUnit.Filter = "UnitID = " & dcUnit.BoundText  'getValueAt("SELECT UnitID,Unit FROM Unit WHERE Unit='" & .TextMatrix(c, 4) & "'", "UnitID")
                
                RSStockUnit!Onhand = RSStockUnit!Onhand + intQtyOld
                
                RSStockUnit.Update
            Else
                Exit Sub
            End If
        End If
        
'        RSStockCard.Filter = "StockID = " & intStockID & " AND RefNo2 = '" & txtRefNo.Text & "'"
'
'        If RSStockCard.RecordCount = 0 Then RSStockCard.AddNew
'
'        'Deduct qty solt to stock card
'        RSStockCard!Type = "A" 'A for assorted product
'        RSStockCard!UnitID = dcUnit.BoundText
'        RSStockCard!RefNo2 = PK
'        RSStockCard!Pieces2 = intQtyOrdered
'        RSStockCard!StockID = intStockID
'
'        RSStockCard.Update
        
        RSStockUnit.Find "UnitID = " & dcUnit.BoundText

        'Deduct qty from highest unit breakdown if Onhand is less than qty ordered
        If RSStockUnit!Onhand < intQtyOrdered Then
            DeductOnhand intQtyOrdered, RSStockUnit!Order, True, RSStockUnit
        End If
        
        'deduct qty from Stock Unit's table
        RSStockUnit.Find "UnitID = " & dcUnit.BoundText
        
        RSStockUnit!Onhand = RSStockUnit!Onhand - intQtyOrdered
        
        RSStockUnit.Update
            
        'Highlight the current row's column
        .ColSel = 3
        'Display a remove button
        If blnAddIncoming = True Then
            intQtyOrdered = intSuggestedQty
            intCount = 2
            GoSub AddIncoming
            
'            blnAddIncoming = False
        End If
        
        Grid_Click
        'Reset the entry fields
        ResetEntry
    End With
    
    Exit Sub
    
GetOnhand:
    intTotalOnhInc = GetTotalQty("Total", RSStockUnit!Order, RSStockUnit!TotalQty, RSStockUnit)
    
    If intTotalOnhInc > 0 Then
    
        intTotalOnhand = GetTotalQty("Onhand", RSStockUnit!Order, RSStockUnit!Onhand, RSStockUnit)
        If intTotalOnhand > 0 Then
        
            If intQtyOrdered > intTotalOnhand Then
                intExcessQty = intQtyOrdered - intTotalOnhand
                
                intTotalIncoming = GetTotalQty("Incoming", RSStockUnit!Order, RSStockUnit!Incoming, RSStockUnit)
                
                If intTotalIncoming > 0 And intTotalIncoming >= intExcessQty Then
                    intSuggestedQty = intExcessQty
                    With frmSuggestedQty
                        .intStockID = intStockID
                        .strProduct = nsdStock.Text
                        .intQtyOrdered = intTotalOnhand
                        .intQtySuggested = intExcessQty
                        
                        .show 1
                            
                        If .blnUseSuggestedQty = True And .blnCancel = False Then
                            blnAddIncoming = True
                            intSuggestedQty = intExcessQty
                        ElseIf .blnCancel = True Then
                            Exit Sub
                        End If
                        
                        intQtyOrdered = intTotalOnhand
                    End With
                Else
                    With frmSuggestedQty
                        .intStockID = intStockID
                        .strProduct = nsdStock.Text
                        .intQtyOrdered = intTotalOnhand
                        .intQtySuggested = intTotalIncoming
                        
                        .show 1
                            
                        If .blnUseSuggestedQty = True And .blnCancel = False Then
                            blnAddIncoming = True
                            intSuggestedQty = intTotalIncoming
                            
                            intCount = 1
                        ElseIf .blnCancel = True Then
                            Exit Sub
                        End If
                        
                        intQtyOrdered = intTotalOnhand
                    End With
                End If
            End If
        End If
    Else
        MsgBox "Insufficient qty", vbInformation
        With frmCustomersItem
            .StockID = intStockID
            
            .show 1
            RSStockUnit.Close
            
            If .blnCancel = False Then
                GoSub GetOnhand
            Else
                Exit Sub
            End If
        End With
    End If
    
    GoSub Continue
erR:
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Function DeductOnhand(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Onhand >= QtyNeededTemp Then
            If blnDeduct = False Then
                DeductOnhand = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Onhand) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Onhand = rs!Onhand - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Onhand = rs!Onhand + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                DeductOnhand = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        DeductOnhand = False
    End If
End Function

'Get the total Qty onhand, incoming and total of onhand and incoming
Private Function GetTotalQty(strField As String, Order As Integer, intOnhand As Integer, rs As Recordset) As Integer
    Dim strFieldValue As Integer
    Dim intOrder As Integer
    
    GetTotalQty = intOnhand
    
    intOrder = Order - 1
    
    Do Until intOrder < 1
        rs.MoveFirst
        rs.Find "Order = " & intOrder
        
        If strField = "Onhand" Then
            strFieldValue = rs!Onhand
        ElseIf strField = "Incoming" Then
            strFieldValue = rs!Incoming
        Else
            strFieldValue = rs!TotalQty
        End If
        
        GetTotalQty = GetTotalQty + GetTotalUnitQty(Order, intOrder, strFieldValue, rs)
        intOrder = intOrder - 1
    Loop
End Function

'This function is called by GetTotalQty Function
Private Function GetTotalUnitQty(Order As Integer, ByVal Ordertmp As Integer, intOnhand As Integer, rs As Recordset)
    GetTotalUnitQty = 1
    Do Until Order = Ordertmp
        Ordertmp = Ordertmp + 1
        
        rs.MoveNext
        
        GetTotalUnitQty = GetTotalUnitQty * rs!Qty
    Loop
    GetTotalUnitQty = intOnhand * GetTotalUnitQty
End Function

Private Function GetIncoming(QtyNeeded As Integer, ByVal Order As Integer, ByVal blnDeduct As Boolean, rs As Recordset) As Boolean
    Dim Onhand As Boolean
    Dim OrderTemp As Integer
    Dim QtyNeededTemp As Double
    
Reloop:
    OrderTemp = Order
    QtyNeededTemp = QtyNeeded
    rs.Find "Order = " & OrderTemp
    
    
    Do Until Onhand = True 'Or OrderTemp = 1
        If rs!Incoming >= QtyNeededTemp Then
            If blnDeduct = False Then
                GetIncoming = True
                Exit Function
            Else
                Onhand = True
            End If
            
            If QtyNeededTemp > 0 And QtyNeededTemp < 1 Then
                QtyNeededTemp = 1
            Else
                QtyNeededTemp = CInt(QtyNeededTemp)
            End If
        Else
            OrderTemp = OrderTemp - 1
            If OrderTemp < 1 Then Exit Do
            QtyNeededTemp = (QtyNeededTemp - rs!Incoming) / rs!Qty
            
            rs.MoveFirst
            
            rs.Find "Order = " & OrderTemp
        End If
    Loop
    
    If Onhand = True Then
        Do
            rs!Incoming = rs!Incoming - QtyNeededTemp
            OrderTemp = OrderTemp + 1
            
            rs.MoveFirst
            rs.Find "Order = " & OrderTemp
            
            rs!Incoming = rs!Incoming + (QtyNeededTemp * rs!Qty)
            
            rs.Update
            
            Onhand = False
            
            If OrderTemp = Order Then
                GetIncoming = True
                Exit Do
            Else
                GoSub Reloop
            End If
        Loop
    Else
        GetIncoming = False
    End If
End Function

Private Sub btnRemove_Click()
    'Remove selected load product
    With Grid
        'Update the record count
        cIRowCount = cIRowCount - 1
        
        If .Rows = 2 Then Grid.Rows = Grid.Rows + 1
        .RemoveItem (.RowSel)
    End With

    'Save to stock card
'    Dim RSStockCard As New Recordset
'
'    With RSStockCard
'        .CursorLocation = adUseClient
'        .Open "SELECT * FROM Stock_Card WHERE StockID = " & toNumber(Grid.TextMatrix(Grid.RowSel, 10)) & " AND RefNo2 = '" & txtRefNo.Text & "'", CN, adOpenStatic, adLockOptimistic
'
'        !Pieces2 = !Pieces2 - toNumber(Grid.TextMatrix(Grid.RowSel, 3))
'
'        .Update
'    End With
'
    btnRemove.Visible = False
    Grid_Click
End Sub

Private Sub CmdTasks_Click()
    PopupMenu mnu_Tasks
End Sub

Private Sub mnu_History_Click()
    On Error Resume Next
    Dim tDate1 As String
    Dim tUser1 As String
    
    tDate1 = Format$(rs.Fields("DateAdded"), "MMM-dd-yyyy HH:MM AMPM")
    
    tUser1 = getValueAt("SELECT PK,CompleteName FROM tbl_SM_Users WHERE PK = " & rs.Fields("AddedByFK"), "CompleteName")
    
    MsgBox "Date Added: " & tDate1 & vbCrLf & _
           "Added By: " & tUser1 & vbCrLf & _
           "" & vbCrLf & _
           "Last Modified: n/a" & vbCrLf & _
           "Modified By: n/a", vbInformation, "Modification History"
           
    tDate1 = vbNullString
    tUser1 = vbNullString
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next

    If blnSave = False Then CN.RollbackTrans
    Unload Me
End Sub

Private Sub cmdSave_Click()
On Error GoTo erR

    'Verify the entries
    If nsdProduct.Text = "" Then
        MsgBox "Please select a product.", vbExclamation
        nsdProduct.SetFocus
        Exit Sub
    End If
   
    If cIRowCount < 1 Then
        MsgBox "Please enter item to purchase before you can save this record.", vbExclamation
        nsdStock.SetFocus
        Exit Sub
    End If
              
    If MsgBox("This save the record. Do you want to proceed?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM Assorted_Product_Detail WHERE AssortedProductID=" & PK, CN, adOpenStatic, adLockOptimistic

    Screen.MousePointer = vbHourglass

    Dim c As Integer

    DeleteItems
    
    'Save the record
    With rs
        If State = adStateAddMode Or State = adStatePopupMode Then
            .AddNew
            ![AssortedProductID] = PK
            ![StockID] = nsdProduct.BoundText
            
            ![DateAdded] = Now
            ![AddedByFK] = CurrUser.USER_PK
        ElseIf State = adStateEditMode Then
            .Close
            .Open "SELECT * FROM Assorted_Product WHERE AssortedProductID=" & PK, CN, adOpenStatic, adLockOptimistic
            
            ![DateModified] = Now
            ![LastUserFK] = CurrUser.USER_PK
        End If
        
        !Qty = txtQtyParent.Text
        !Date = dtpDate.Value
        ![Status] = IIf(cboStatus.Text = "Assorted", True, False)
        ![Notes] = txtNotes.Text

        .Update
    End With
    
    Dim intUnitsOrder As Integer
    Dim intQty As Integer
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            If State = adStateAddMode Or State = adStatePopupMode Then
AddNew:
                'Add qty received in Local Purchase Details
                RSDetails.AddNew

                RSDetails![AssortedProductID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 4))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 2))
                RSDetails![UnitID] = getUnitID(.TextMatrix(c, 3))
                
                RSDetails.Update
                
            ElseIf State = adStateEditMode Then
                RSDetails.Filter = "StockID = " & toNumber(.TextMatrix(c, 4))
            
                If RSDetails.RecordCount = 0 Then GoTo AddNew
                
                RSDetails![AssortedProductID] = PK
                RSDetails![StockID] = toNumber(.TextMatrix(c, 4))
                RSDetails![Qty] = toNumber(.TextMatrix(c, 2))
                RSDetails![UnitID] = getUnitID(.TextMatrix(c, 3))
                
                RSDetails.Update
                
            End If
            
        Next c
    End With
    
    'Clear variables
    c = 0
    Set RSDetails = Nothing

    CN.CommitTrans
    
    blnSave = True
    
    HaveAction = True
    Screen.MousePointer = vbDefault

    If State = adStateAddMode Then
        MsgBox "New record has been successfully saved.", vbInformation
        
        If MsgBox("Do you want to add another new record?", vbQuestion + vbYesNo) = vbYes Then
            ResetFields
            GeneratePK
        Else
            Unload Me
        End If
    Else
        MsgBox "Changes in record has been successfully saved.", vbInformation
        Unload Me
    End If

    Exit Sub
erR:
    blnSave = False
'    CN.RollbackTrans
'    CN.BeginTrans
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    If CloseMe = True Then
        Unload Me
    Else
        nsdProduct.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then SendKeys ("{tab}")
End Sub

Private Sub Form_Load()
    Dim strRoute As String
    
    InitGrid
    
    'Check the form state
    If State = adStateAddMode Or State = adStatePopupMode Then
        InitNSD
    
    'Set the recordset
    If rs.State = 1 Then rs.Close
        rs.Open "SELECT * FROM Assorted_Product WHERE AssortedProductID=" & PK, CN, adOpenStatic, adLockOptimistic
        dtpDate.Value = Date
        
        CN.BeginTrans

        GeneratePK
    Else
        Screen.MousePointer = vbHourglass
        'Set the recordset
        rs.Open "SELECT * FROM qry_Assorted_Product WHERE AssortedProductID=" & PK, CN, adOpenStatic, adLockOptimistic
        
        If State = adStateViewMode Then
            cmdCancel.Caption = "Close"
                   
            DisplayForViewing
        Else
            InitNSD
            
            CN.BeginTrans
            
            DisplayForEditing
        End If
    
        If ForCusAcc = True Then
            Me.Icon = frmSalesReceipts.Icon
        End If

        Screen.MousePointer = vbDefault
    End If
    
    'Initialize Graphics
    With MAIN
        'cmdGenerate.Picture = .i16x16.ListImages(14).Picture
        'cmdNew.Picture = .i16x16.ListImages(10).Picture
        'cmdReset.Picture = .i16x16.ListImages(15).Picture
    End With
End Sub

'Procedure used to generate PK
Private Sub GeneratePK()
    PK = getIndex("Assorted_Product")
End Sub

Private Sub ResetEntry()
    nsdStock.ResetValue
    txtQty.Text = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If HaveAction = True Then
        frmAssortedProduct.RefreshRecords
    End If
    
    Set frmAssortedProductAE = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        If State = adStateViewMode Then Exit Sub

        nsdStock.Text = .TextMatrix(.RowSel, 1)
        nsdStock.Tag = .TextMatrix(.RowSel, 4) 'Add tag coz boundtext is empty
        intQtyOld = IIf(.TextMatrix(.RowSel, 2) = "", 0, .TextMatrix(.RowSel, 2))
        txtQty.Text = .TextMatrix(.RowSel, 2)
        
        On Error Resume Next
        bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & .TextMatrix(.RowSel, 4), "Unit", dcUnit, "UnitID", True
        On Error GoTo 0
        
        dcUnit.Text = .TextMatrix(.RowSel, 3)
        'disable unit to prevent user from changing it. changing of unit will result to imbalance of inventory
        If State = adStateEditMode Then dcUnit.Enabled = False
    
        If Grid.Rows = 2 And Grid.TextMatrix(1, 4) = "" Then '4 = StockID
            btnRemove.Visible = False
        Else
            btnRemove.Visible = True
            btnRemove.Top = (Grid.CellTop + Grid.Top) - 20
            btnRemove.Left = Grid.Left + 50
        End If
    End With
End Sub

Private Sub Grid_Scroll()
    btnRemove.Visible = False
End Sub

Private Sub Grid_SelChange()
    Grid_Click
End Sub


Private Sub nsdStock_Change()
    On Error Resume Next
    
    nsdStock.Tag = nsdStock.BoundText
    txtQty.Text = "0"
        
    dcUnit.Text = ""
    bind_dc "SELECT * FROM qry_Unit WHERE StockID=" & nsdStock.BoundText, "Unit", dcUnit, "UnitID", True
    
'    txtPrice.Text = toMoney(nsdStock.getSelValueAt(3)) 'Selling Price
End Sub

Private Sub txtQty_Validate(Cancel As Boolean)
    txtQty.Text = toNumber(txtQty.Text)
End Sub

Private Sub txtQty_Change()
    If toNumber(txtQty.Text) < 1 Then
        btnAdd.Enabled = False
        Exit Sub
    Else
        btnAdd.Enabled = True
    End If
End Sub

Private Sub txtQty_GotFocus()
    HLText txtQty
End Sub

'Procedure used to reset fields
Private Sub ResetFields()
    InitGrid
    ResetEntry
    
'    nsdClient.Text = ""
'    txtSONo.Text = ""
'    txtAddress.Text = ""
'    txtDate.Text = ""
'    txtSalesman.Text = ""
'    txtOrderedBy.Text = ""
'    txtDispatchedBy.Text = ""
'    txtDeliveryInstructions.Text = ""
'
'    txtGross(2).Text = "0.00"
'    txtDesc.Text = "0.00"
'    txtTaxBase.Text = "0.00"
'    txtVat.Text = "0.00"
'    txtNet.Text = "0.00"
'
'    cIAmount = 0
'    cDAmount = 0
'
'    nsdClient.SetFocus
End Sub

'Used to display record
Private Sub DisplayForEditing()
    On Error GoTo erR
    
    nsdProduct.Tag = rs!StockID
    nsdProduct.DisableDropdown = True
    nsdProduct.TextReadOnly = True
    nsdProduct.Text = rs![Stock]
    txtQtyParent.Text = rs![Qty]
    dtpDate.Value = rs![Date]
    cboStatus.Text = rs!Status_Alias
    txtNotes.Text = rs![Notes]
    
    cIRowCount = 0

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Assorted_Product_Detail WHERE AssortedProductID=" & PK & " ORDER BY AProdDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            cIRowCount = cIRowCount + 1     'increment
            With Grid
                If .Rows = 2 And .TextMatrix(1, 4) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Stock]
                    .TextMatrix(1, 2) = RSDetails![Qty]
                    .TextMatrix(1, 3) = RSDetails![Unit]
                    .TextMatrix(1, 4) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 3
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionByRow
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    cmdSave.Caption = "Save"
    
    Exit Sub
erR:
    'Error if encounter a null value
    If erR.Number = 94 Then
        Resume Next
    Else
        MsgBox erR.Description
    End If
End Sub

'Used to display record
Private Sub DisplayForViewing()
    On Error GoTo erR
    
    nsdProduct.Tag = rs!StockID
    nsdProduct.DisableDropdown = True
    nsdProduct.TextReadOnly = True
    nsdProduct.Text = rs![Stock]
    txtQtyParent.Text = rs![Qty]
    txtDate.Text = rs![Date]
    cboStatus.Text = rs!Status_Alias
    txtNotes.Text = rs![Notes]

    'Display the details
    Dim RSDetails As New Recordset

    RSDetails.CursorLocation = adUseClient
    RSDetails.Open "SELECT * FROM qry_Assorted_Product_Detail WHERE AssortedProductID=" & PK & " ORDER BY AProdDetailID ASC", CN, adOpenStatic, adLockOptimistic
    If RSDetails.RecordCount > 0 Then
        RSDetails.MoveFirst
        While Not RSDetails.EOF
            With Grid
                If .Rows = 2 And .TextMatrix(1, 4) = "" Then
                    .TextMatrix(1, 1) = RSDetails![Stock]
                    .TextMatrix(1, 2) = RSDetails![Qty]
                    .TextMatrix(1, 3) = RSDetails![Unit]
                    .TextMatrix(1, 4) = RSDetails![StockID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = RSDetails![Stock]
                    .TextMatrix(.Rows - 1, 2) = RSDetails![Qty]
                    .TextMatrix(.Rows - 1, 3) = RSDetails![Unit]
                    .TextMatrix(.Rows - 1, 4) = RSDetails![StockID]
                End If
            End With
            RSDetails.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 4
        'Set fixed cols
        If State = adStateEditMode Then
            Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionByRow
            Grid.FixedCols = 1
        End If
    End If

    RSDetails.Close
    'Clear variables
    Set RSDetails = Nothing

    'Disable commands
    LockInput Me, True

    dtpDate.Visible = False
    txtDate.Visible = True
    picPurchase.Visible = False
    cmdSave.Visible = False
    btnAdd.Visible = False

    'Resize and reposition the controls
   
    Shape3.Top = 1850
    Label11.Top = 1850
    Line1(1).Visible = False
    Line2(1).Visible = False
    Grid.Top = 2200
    Grid.Height = 3280

    Exit Sub
erR:
    'Error if encounter a null value
    If erR.Number = 94 Then
        Resume Next
    Else
        MsgBox erR.Description
    End If
End Sub

'Procedure used to initialize the grid
Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 5
        .ColSel = 4
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 4000
        .ColWidth(2) = 1505
        .ColWidth(3) = 1545
        .ColWidth(4) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Description"
        .TextMatrix(0, 2) = "Qty"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Stock ID"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
'        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
    End With
End Sub

Private Sub InitNSD()
    'For Stock
    With nsdProduct
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,StockID"
        .sqlTables = "Stocks"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Products"
    End With
    
    'For Stock
    With nsdStock
        .ClearColumn
        .AddColumn "Barcode", 2064.882
        .AddColumn "Stock", 4085.26
        
        .Connection = CN.ConnectionString
        
        .sqlFields = "Barcode,Stock,StockID"
        .sqlTables = "Stocks"
        .sqlSortOrder = "Stock ASC"
        .BoundField = "StockID"
        .PageBy = 25
        .DisplayCol = 2
        
        .setDropWindowSize 6800, 4000
        .TextReadOnly = True
        .SetDropDownTitle = "Products"
    End With
End Sub

Private Sub DeleteItems()
    Dim CurrRow As Integer
    Dim RSAssorted As New Recordset

    If State = adStateAddMode Then Exit Sub

    RSAssorted.CursorLocation = adUseClient
    RSAssorted.Open "SELECT * FROM Assorted_Product_Detail WHERE AssortedProductID=" & PK, CN, adOpenStatic, adLockOptimistic
    If RSAssorted.RecordCount > 0 Then
        RSAssorted.MoveFirst
        While Not RSAssorted.EOF
            CurrRow = getFlexPos(Grid, 4, RSAssorted!StockID)

            'Add to grid
            With Grid
                If CurrRow < 0 Then
                    'Delete record if doesnt exist in flexgrid
                    DelRecwSQL "Assorted_Product_Detail", "AProdDetailID", "", True, RSAssorted!AProdDetailID
                End If
            End With
            RSAssorted.MoveNext
        Wend
    End If
End Sub
