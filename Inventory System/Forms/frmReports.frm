VERSION 5.00
Object = "{C4847593-972C-11D0-9567-00A0C9273C2A}#8.0#0"; "crviewer.dll"
Begin VB.Form frmReports 
   Caption         =   "Form1"
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   ScaleHeight     =   6645
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin CRVIEWERLibCtl.CRViewer CR 
      Height          =   3915
      Left            =   930
      TabIndex        =   0
      Top             =   1200
      Width           =   6405
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
   End
End
Attribute VB_Name = "frmReports"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public strReport        As String
Public PK               As String
Public strYear          As String
Public blnPaid          As Boolean
Public strWhere         As String

Dim mTest As CRAXDRT.Application
Dim mReport As CRAXDRT.Report
Dim SubReport As CRAXDRT.Report
Dim mParam As CRAXDRT.ParameterFieldDefinitions

Public Sub CommandPass(ByVal srcPerformWhat As String)
    Select Case srcPerformWhat
        Case "Close"
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
On Error GoTo err_Form_Load
    Dim mSubRep
    
    Set mTest = New CRAXDRT.Application
    Set mReport = New CRAXDRT.Report
    
    Select Case strReport
        Case "Customers Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Customers_Report.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Customers Profile Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Customers_Report.rpt")
                      
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Expenses Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Expenses_Report.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Loading Form Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Loading_Form.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Purchase Returns Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Purchase_Returns.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Purchase Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Purchases_Report.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Receipt Form Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Receipt_Form.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Sales Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Sales_Report.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Sales Return Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Sales_Returns.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        Case "Suppliers Report"
            Set mReport = mTest.OpenReport(App.Path & "\Reports\rpt_Suppliers_Report.rpt")
            
            mReport.RecordSelectionFormula = strWhere
            
            Set mParam = mReport.ParameterFields
            mParam.Item(1).AddCurrentValue CurrBiz.BUSINESS_NAME
        End Select
    
    Screen.MousePointer = vbHourglass
    CR.ReportSource = mReport
    CR.ViewReport
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
err_Form_Load:
    prompt_err err, Name, "Form_Load"
    Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Resize()
    With CR
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmReports = Nothing
End Sub

