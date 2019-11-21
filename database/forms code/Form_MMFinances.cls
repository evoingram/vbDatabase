VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_MMFinances"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private Sub btnViewReport_Click()
'============================================================================
' Name        : btnViewReport_Click
' Author      : Erica L Ingram
' Copyright   : 2019, A Quo Co.
' Call command: Call btnViewReport_Click
' Description : generates document from combo box selection
'============================================================================


If Me.cmbListAllDocuments = "List of Approved Customers" Then
    Me.YearlySalesReport.SourceObject = "Report.ApprovedCustomers"
    
ElseIf Me.cmbListAllDocuments = "Single Invoice (need job number)" Then
    Forms![MMFinances]![YearlySalesReport].SourceObject = "ViewInvoice"
    
ElseIf Me.cmbListAllDocuments = "Customer History by Year" Then
    Me.YearlySalesReport.SourceObject = "Report.Customer History by Year"
        
ElseIf Me.cmbListAllDocuments = "Incomplete Tasks" Then
    Me.YearlySalesReport.SourceObject = "Report.GroupTasksIncompleteProduction1"

ElseIf Me.cmbListAllDocuments = "Tasks Scheduled" Then
    Me.YearlySalesReport.SourceObject = "Report.Tasks Scheduled"

ElseIf Me.cmbListAllDocuments = "Unpaid Invoices -- Detailed" Then
    Me.YearlySalesReport.SourceObject = "Report.QBalanceOwed"

ElseIf Me.cmbListAllDocuments = "Unpaid Invoices" Then
    Me.YearlySalesReport.SourceObject = "Report.UnpaidInvoices"


ElseIf Me.cmbListAllDocuments = "Unpaid PayPal Invoices" Then
    Me.YearlySalesReport.SourceObject = "Report.UnpaidPPInvoices"
    
ElseIf Me.cmbListAllDocuments = "Profit/Loss Income Report" Then
    Me.YearlySalesReport.SourceObject = "Report.ProfitLossIncome"
ElseIf Me.cmbListAllDocuments = "Sales Report -- Yearly" Then
    Me.YearlySalesReport.SourceObject = "Report.YearlySalesReport"

ElseIf Me.cmbListAllDocuments = "Sales Report -- Quarterly" Then
    Me.YearlySalesReport.SourceObject = "Report.QuarterlySalesReport"

ElseIf Me.cmbListAllDocuments = "Sales Report -- Monthly" Then
    Me.YearlySalesReport.SourceObject = "Report.SalesReportMonthly"
 
ElseIf Me.cmbListAllDocuments = "Sales Report -- Monthly Projected" Then
    Me.YearlySalesReport.SourceObject = "Report.SalesReportMonthlyProjected"
               
ElseIf Me.cmbListAllDocuments = "Sales Report -- Monthly Detailed" Then
    Me.YearlySalesReport.SourceObject = "Report.MonthlySalesDetailedReport"

ElseIf Me.cmbListAllDocuments = "Sales Report -- Monthly Invoice Detailed" Then
    Me.YearlySalesReport.SourceObject = "Report.SalesReportMonthlyInvoiceDetailed"

ElseIf Me.cmbListAllDocuments = "Sales Report -- Month over Month" Then
    Me.YearlySalesReport.SourceObject = "Report.SalesReportMonthOverMonth"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Monthly by Category" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportMonthlybyCategory"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Monthly" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportMonthly"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Quarterly" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportQuarterly"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Yearly" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportYearly"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Yearly by Category" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportYearlybyCategory"

ElseIf Me.cmbListAllDocuments = "Expenses Report -- Quarterly by Category" Then
    Me.YearlySalesReport.SourceObject = "Report.ExpensesReportQuarterlybyCategory"
Else
End If

End Sub

Private Sub Command73_Click()

Forms![NewMainMenu]![ProcessJobSubformNMM].SourceObject = "fPayments"
Forms![NewMainMenu]![ProcessJobSubformNMM].Requery
End Sub


