VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmLandedCost 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Landed Cost"
   ClientHeight    =   7635
   ClientLeft      =   645
   ClientTop       =   2775
   ClientWidth     =   15210
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   15210
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "Update"
      Height          =   315
      Left            =   14340
      TabIndex        =   40
      Top             =   900
      Width           =   765
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Default         =   -1  'True
      Height          =   315
      Left            =   12390
      TabIndex        =   39
      Top             =   7200
      Width           =   1275
   End
   Begin VB.TextBox txtRefNo 
      Height          =   315
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1260
      Width           =   675
   End
   Begin VB.TextBox txtMarkup 
      Height          =   315
      Left            =   10680
      TabIndex        =   13
      Text            =   "0.00"
      Top             =   1260
      Width           =   735
   End
   Begin VB.TextBox txtTelNo 
      Height          =   315
      Left            =   8640
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   180
      Width           =   1485
   End
   Begin VB.Frame Frame2 
      Height          =   60
      Left            =   150
      TabIndex        =   35
      Top             =   540
      Width           =   14985
   End
   Begin VB.TextBox txtLocation 
      Height          =   315
      Left            =   6150
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   180
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   60
      TabIndex        =   33
      Top             =   7020
      Width           =   15015
   End
   Begin VB.CommandButton cmdForwardPrice 
      Caption         =   "&Forward selected item's price to products profile"
      Height          =   315
      Left            =   8610
      TabIndex        =   18
      Top             =   7200
      Width           =   3705
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   315
      Left            =   13770
      TabIndex        =   19
      Top             =   7200
      Width           =   1335
   End
   Begin VB.TextBox txtNotes 
      Height          =   315
      Left            =   13140
      TabIndex        =   16
      Top             =   1260
      Width           =   1995
   End
   Begin VB.TextBox txtExtensionPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   12330
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   1260
      Width           =   765
   End
   Begin VB.TextBox txtSellingPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   11490
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   1260
      Width           =   765
   End
   Begin VB.TextBox txtFreight 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9870
      TabIndex        =   12
      Text            =   "0.00"
      Top             =   1260
      Width           =   765
   End
   Begin VB.TextBox txtDiscountedSP 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9000
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox txtExtDiscAmt 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   8070
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "0.00"
      Top             =   1260
      Width           =   885
   End
   Begin VB.TextBox txtExtDiscPerc 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   7200
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "0"
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox txtDiscount 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   6330
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "0"
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox txtSupplierPrice 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   5520
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "0.00"
      Top             =   1260
      Width           =   765
   End
   Begin VB.TextBox txtDate 
      Height          =   315
      Left            =   4470
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   1260
      Width           =   1005
   End
   Begin VB.TextBox txtPackaging 
      Height          =   315
      Left            =   3600
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1260
      Width           =   825
   End
   Begin VB.TextBox txtProduct 
      Height          =   315
      Left            =   840
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1260
      Width           =   2715
   End
   Begin VB.TextBox txtSupplier 
      Height          =   315
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   180
      Width           =   4155
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid 
      Height          =   5160
      Left            =   90
      TabIndex        =   17
      Top             =   1680
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   9102
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
      AllowUserResizing=   1
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
   Begin VB.Label Label17 
      Caption         =   "Ref No"
      Height          =   225
      Left            =   120
      TabIndex        =   38
      Top             =   1020
      Width           =   645
   End
   Begin VB.Label Label16 
      Caption         =   "Markup"
      Height          =   225
      Left            =   10710
      TabIndex        =   37
      Top             =   1020
      Width           =   735
   End
   Begin VB.Label Label15 
      Caption         =   "Tel. No."
      Height          =   285
      Left            =   7770
      TabIndex        =   36
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label Label14 
      Caption         =   "Location:"
      Height          =   285
      Left            =   5280
      TabIndex        =   34
      Top             =   210
      Width           =   1305
   End
   Begin VB.Label Label13 
      Caption         =   "Note"
      Height          =   195
      Left            =   13140
      TabIndex        =   32
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      Caption         =   "Extension Price "
      Height          =   405
      Left            =   12330
      TabIndex        =   31
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   "Selling Price "
      Height          =   375
      Left            =   11490
      TabIndex        =   30
      Top             =   840
      Width           =   735
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      Caption         =   "Freight"
      Height          =   195
      Left            =   9870
      TabIndex        =   29
      Top             =   1050
      Width           =   735
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount Supplier Price "
      Height          =   615
      Left            =   9000
      TabIndex        =   28
      Top             =   600
      Width           =   795
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Ext. Discount Amt "
      Height          =   585
      Left            =   8070
      TabIndex        =   27
      Top             =   630
      Width           =   885
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Ext. Discount % "
      Height          =   375
      Left            =   7200
      TabIndex        =   26
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Discount"
      Height          =   195
      Left            =   6330
      TabIndex        =   25
      Top             =   1020
      Width           =   795
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Supplier Price "
      Height          =   405
      Left            =   5520
      TabIndex        =   24
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label4 
      Caption         =   "Date"
      Height          =   195
      Left            =   4470
      TabIndex        =   23
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label3 
      Caption         =   "Packaging"
      Height          =   195
      Left            =   3600
      TabIndex        =   22
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label2 
      Caption         =   "Item/ Product"
      Height          =   195
      Left            =   840
      TabIndex        =   21
      Top             =   1020
      Width           =   1035
   End
   Begin VB.Label Label1 
      Caption         =   "Supplier:"
      Height          =   285
      Left            =   120
      TabIndex        =   20
      Top             =   210
      Width           =   1305
   End
End
Attribute VB_Name = "frmLandedCost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intProductID         As String
Public intSupplierID        As String
Dim rs                      As New Recordset
Dim cIRowCount              As Integer

Private Sub CmdClose_Click()
    rs.Close
    Set rs = Nothing
    
    Unload Me
End Sub

Private Sub cmdForwardPrice_Click()
    Dim RSProduct As New Recordset
    Dim RSProductUnit As New Recordset
    Dim cSalesPrice As Currency
    
    CN.BeginTrans
    
    With RSProduct
        .CursorLocation = adUseClient
        .Open "SELECT StockID, ExtPrice FROM Stocks WHERE StockID=" & intProductID, CN, adOpenStatic, adLockOptimistic
    
        If .RecordCount > 0 Then
            !ExtPrice = txtExtensionPrice.Text
            
            .Update
        Else
            MsgBox "Oops! Product ID " & intProductID & " does not exist", vbCritical
            
            CN.RollbackTrans
            
            Exit Sub
        End If
    End With
    
    cSalesPrice = txtSellingPrice.Text

    With RSProductUnit
    
        .CursorLocation = adUseClient
        .Open "SELECT StockUnitID, Order, Qty, SalesPrice FROM qry_Stock_Unit WHERE StockID=" & intProductID & " ORDER BY qry_Stock_Unit.[Order] ASC", CN, adOpenStatic, adLockOptimistic
        
        If .RecordCount > 0 Then
            !SalesPrice = cSalesPrice
            .MoveNext
            Do While Not .EOF
                cSalesPrice = cSalesPrice / !Qty
                
                !SalesPrice = cSalesPrice
                .Update
                .MoveNext
            Loop
        End If
    End With
    
    CN.CommitTrans
    
    RSProduct.Close
    Set RSProduct = Nothing
    
    Exit Sub
    
erR:
    CN.RollbackTrans
    prompt_err erR, Name, "cmdForwardPrice_Click"
End Sub

Private Sub CmdSave_Click()
    Dim c As Integer
      
    CN.BeginTrans
    
    With Grid
        'Save the details of the records
        For c = 1 To cIRowCount
            .Row = c
            rs.MoveFirst
            rs.Find "LandedCostID = " & toNumber(.TextMatrix(c, 15))
        
            rs![RefNo] = .TextMatrix(c, 1)
            rs![Unit] = .TextMatrix(c, 3)
            rs![Date] = .TextMatrix(c, 4)
            rs![SupplierPrice] = .TextMatrix(c, 5)
            rs![Discount] = .TextMatrix(c, 6) / 100
            rs![ExtDiscPercent] = .TextMatrix(c, 7) / 100
            rs![ExtDiscAmount] = .TextMatrix(c, 8)
            rs![DiscountedSP] = .TextMatrix(c, 9)
            rs![Freight] = .TextMatrix(c, 10)
            rs![Markup] = .TextMatrix(c, 11)
            rs![SellingPrice] = .TextMatrix(c, 12)
            rs![ExtensionPrice] = .TextMatrix(c, 13)
            rs![Notes] = .TextMatrix(c, 14)
            
            rs.Update
        Next c
    End With

    'Clear variables
    c = 0

    CN.CommitTrans

    Screen.MousePointer = vbDefault

    MsgBox "Changes in record has been successfully saved.", vbInformation

    Exit Sub
erR:
    CN.RollbackTrans
    prompt_err erR, Name, "cmdSave_Click"
    Screen.MousePointer = vbDefault
End Sub

Private Sub CmdUpdate_Click()
    Dim CurrRow As Integer
    
    CurrRow = getFlexPos(Grid, 15, txtRefNo.Tag)

    'Update to grid
    With Grid
        .Row = CurrRow
               
'        .TextMatrix(CurrRow, 1) = txtRefNo.Text
'        .TextMatrix(CurrRow, 2) = txtProduct.Text
'        .TextMatrix(CurrRow, 3) = txtPackaging.Text
'        .TextMatrix(CurrRow, 4) = txtDate.Text
'        .TextMatrix(CurrRow, 5) = toMoney(txtSupplierPrice.Text)
'        .TextMatrix(CurrRow, 6) = toNumber(txtDiscount.Text)
'        .TextMatrix(CurrRow, 7) = toNumber(txtExtDiscPerc.Text)
'        .TextMatrix(CurrRow, 8) = toMoney(txtExtDiscAmt.Text)
'        .TextMatrix(CurrRow, 9) = toMoney(DiscountedSP.Text)
        .TextMatrix(CurrRow, 10) = toMoney(txtFreight.Text)
        .TextMatrix(CurrRow, 11) = toMoney(txtMarkup.Text)
        .TextMatrix(CurrRow, 12) = toMoney(txtSellingPrice.Text)
        .TextMatrix(CurrRow, 13) = toMoney(txtExtensionPrice.Text)
        .TextMatrix(CurrRow, 14) = txtNotes.Text

        'Highlight the current row's column
        .ColSel = 14
        'Display a remove button
        Grid_Click
    End With
End Sub

Private Sub Form_Load()
    InitGrid
    
    rs.Open "SELECT * FROM qry_Landed_Cost WHERE StockID=" & intProductID & " AND VendorID=" & intSupplierID, CN, adOpenStatic, adLockOptimistic
    
    DisplayForEditing
End Sub

Private Sub DisplayForEditing()
    Dim cDiscountedSP As Currency 'Discounted Supplier Price
    
    If rs.RecordCount > 0 Then
    txtSupplier.Text = rs!Company
    
        rs.MoveFirst
        While Not rs.EOF
            With Grid
                cIRowCount = cIRowCount + 1     'increment
                If .Rows = 2 And .TextMatrix(1, 15) = "" Then
                    .TextMatrix(1, 1) = rs![RefNo]
                    .TextMatrix(1, 2) = rs![Stock]
                    .TextMatrix(1, 3) = rs![Unit]
                    .TextMatrix(1, 4) = rs![Date]
                    .TextMatrix(1, 5) = toMoney(rs![SupplierPrice])
                    .TextMatrix(1, 6) = rs![Discount] * 100
                    .TextMatrix(1, 7) = rs![ExtDiscPercent] * 100
                    .TextMatrix(1, 8) = toMoney(rs![ExtDiscAmount])
                    cDiscountedSP = toMoney(rs![SupplierPrice]) - (rs![Discount] * 100) - (rs![ExtDiscPercent] * 100) - toMoney(rs![ExtDiscAmount])
                    .TextMatrix(1, 9) = toMoney(cDiscountedSP)
                    .TextMatrix(1, 10) = toMoney(rs![Freight])
                    .TextMatrix(1, 11) = toMoney(rs![Markup])
                    .TextMatrix(1, 12) = toMoney(rs![SellingPrice])
                    .TextMatrix(1, 13) = toMoney(rs![ExtensionPrice])
                    .TextMatrix(1, 14) = rs!Notes
                    .TextMatrix(1, 15) = rs![LandedCostID]
                Else
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 1) = rs![RefNo]
                    .TextMatrix(.Rows - 1, 2) = rs![Stock]
                    .TextMatrix(.Rows - 1, 3) = rs![Unit]
                    .TextMatrix(.Rows - 1, 4) = rs![Date]
                    .TextMatrix(.Rows - 1, 5) = toMoney(rs![SupplierPrice])
                    .TextMatrix(.Rows - 1, 6) = rs![Discount] * 100
                    .TextMatrix(.Rows - 1, 7) = rs![ExtDiscPercent] * 100
                    .TextMatrix(.Rows - 1, 8) = toMoney(rs![ExtDiscAmount])
                    cDiscountedSP = toMoney(rs![SupplierPrice]) - (rs![Discount] * 100) - (rs![ExtDiscPercent] * 100) - toMoney(rs![ExtDiscAmount])
                    .TextMatrix(.Rows - 1, 9) = toMoney(cDiscountedSP)
                    .TextMatrix(.Rows - 1, 10) = toMoney(rs![Freight])
                    .TextMatrix(.Rows - 1, 11) = toMoney(rs![Markup])
                    .TextMatrix(.Rows - 1, 12) = toMoney(rs![SellingPrice])
                    .TextMatrix(.Rows - 1, 13) = toMoney(rs![ExtensionPrice])
                    .TextMatrix(.Rows - 1, 14) = toMoney(rs![Notes])
                    .TextMatrix(.Rows - 1, 15) = rs![LandedCostID]
                End If
            End With
            rs.MoveNext
        Wend
        Grid.Row = 1
        Grid.ColSel = 14
        'Set fixed cols
        Grid.FixedRows = Grid.Row: 'Grid.SelectionMode = flexSelectionFree
        Grid.FixedCols = 2
    End If
End Sub

Private Sub InitGrid()
    cIRowCount = 0
    With Grid
        .Clear
        .ClearStructure
        .Rows = 2
        .FixedRows = 1
        .FixedCols = 1
        .Cols = 16
        .ColSel = 15
        'Initialize the column size
        .ColWidth(0) = 315
        .ColWidth(1) = 1000
        .ColWidth(2) = 2430
        .ColWidth(3) = 465
        .ColWidth(4) = 1000
        .ColWidth(5) = 1005
        .ColWidth(6) = 900
        .ColWidth(7) = 690
        .ColWidth(8) = 995
        .ColWidth(9) = 1150
        .ColWidth(10) = 1000
        .ColWidth(11) = 1000
        .ColWidth(12) = 1000
        .ColWidth(13) = 1000
        .ColWidth(14) = 1000
        .ColWidth(15) = 0
        'Initialize the column name
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "Ref No"
        .TextMatrix(0, 2) = "Stock"
        .TextMatrix(0, 3) = "Unit"
        .TextMatrix(0, 4) = "Date"
        .TextMatrix(0, 5) = "Supplier Price" 'Supplier Price
        .TextMatrix(0, 6) = "Disc(%)"
        .TextMatrix(0, 7) = "Ext. Disc(%)"
        .TextMatrix(0, 8) = "Ext. Disc(Amt)"
        .TextMatrix(0, 9) = "DiscountedSP"
        .TextMatrix(0, 10) = "Freight"
        .TextMatrix(0, 11) = "Markup"
        .TextMatrix(0, 12) = "Selling Price"
        .TextMatrix(0, 13) = "Ext. Price"
        .TextMatrix(0, 14) = "Notes"
        .TextMatrix(0, 15) = "Landed Cost ID"
        'Set the column alignment
'        .ColAlignment(0) = vbLeftJustify
        .ColAlignment(1) = vbLeftJustify
        .ColAlignment(2) = vbLeftJustify
'        .ColAlignment(3) = vbLeftJustify
'        .ColAlignment(4) = vbRightJustify
'        .ColAlignment(5) = vbLeftJustify
'        .ColAlignment(6) = vbRightJustify
'        .ColAlignment(7) = vbRightJustify
'        .ColAlignment(8) = vbRightJustify
'        .ColAlignment(9) = vbRightJustify
'        .ColAlignment(10) = vbRightJustify
'        .ColAlignment(11) = vbRightJustify
'        .ColAlignment(12) = vbRightJustify
'        .ColAlignment(13) = vbRightJustify
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmLandedCost = Nothing
End Sub

Private Sub Grid_Click()
    With Grid
        txtRefNo.Tag = .TextMatrix(.RowSel, 15)
        txtRefNo.Text = .TextMatrix(.RowSel, 1)
        txtProduct.Text = .TextMatrix(.RowSel, 2)
        txtPackaging.Text = .TextMatrix(.RowSel, 3)
        txtDate.Text = .TextMatrix(.RowSel, 4)
        txtSupplierPrice.Text = .TextMatrix(.RowSel, 5)
        txtDiscount.Text = .TextMatrix(.RowSel, 6)
        txtExtDiscPerc.Text = .TextMatrix(.RowSel, 7)
        txtExtDiscAmt.Text = .TextMatrix(.RowSel, 8)
        txtDiscountedSP.Text = .TextMatrix(.RowSel, 9)
        txtFreight.Text = .TextMatrix(.RowSel, 10)
        txtMarkup.Text = .TextMatrix(.RowSel, 11)
        txtSellingPrice.Text = .TextMatrix(.RowSel, 12)
        txtExtensionPrice.Text = .TextMatrix(.RowSel, 13)
        txtNotes.Text = .TextMatrix(.RowSel, 14)
    End With
End Sub

Private Sub txtExtensionPrice_GotFocus()
    HLText txtExtensionPrice
End Sub

Private Sub txtExtensionPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtExtensionPrice_Validate(Cancel As Boolean)
    txtExtensionPrice.Text = toMoney(toNumber(txtExtensionPrice.Text))
End Sub

Private Sub txtFreight_Change()
    txtSellingPrice = toMoney(txtDiscountedSP.Text) + toMoney(txtFreight.Text) + toMoney(txtMarkup.Text)
End Sub

Private Sub txtFreight_GotFocus()
    HLText txtFreight
End Sub

Private Sub txtFreight_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtFreight_Validate(Cancel As Boolean)
    txtFreight.Text = toMoney(toNumber(txtFreight.Text))
End Sub

Private Sub txtMarkup_Change()
    txtSellingPrice = toMoney(txtDiscountedSP.Text) + toMoney(txtFreight.Text) + toMoney(txtMarkup.Text)
End Sub

Private Sub txtMarkup_GotFocus()
    HLText txtMarkup
End Sub

Private Sub txtMarkup_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtMarkup_Validate(Cancel As Boolean)
    txtMarkup.Text = toMoney(toNumber(txtMarkup.Text))
End Sub

Private Sub txtSellingPrice_GotFocus()
    HLText txtSellingPrice
End Sub

Private Sub txtSellingPrice_KeyPress(KeyAscii As Integer)
    KeyAscii = isNumber(KeyAscii)
End Sub

Private Sub txtSellingPrice_Validate(Cancel As Boolean)
    txtSellingPrice.Text = toMoney(toNumber(txtSellingPrice.Text))
End Sub
