SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Cases].[ID], [Cases].[Party1], [Cases].[Party2], [Cases].[CaseNumber1], [Cases].[Jurisdiction]
FROM Cases
ORDER BY [ID];

SELECT [Customers].[ID], [Customers].[Company], [Customers].[LastName], [Customers].[FirstName], [Customers].[Address], [Customers].[City]
FROM Customers
ORDER BY [LastName];

SELECT [Statuses].[ID], [Statuses].[CourtDatesID]
FROM Statuses;

SELECT [MailClass].[ID], [MailClass].[MailClass], [MailClass].[Description1]
FROM MailClass
ORDER BY [ID];

SELECT [PackageType].[ID], [PackageType].[PackageType], [PackageType].[Description1]
FROM PackageType
ORDER BY [ID];

PARAMETERS __InvoiceNo Value;
SELECT DISTINCTROW *
FROM QInfobyInvoiceNumber AS [INV-SBFM-InvoiceEstmPriceQuote]
WHERE ([__InvoiceNo] = InvoiceNo);

PARAMETERS __OrderingID Value;
SELECT DISTINCTROW *
FROM OrderingAttorneyInfo AS [INV-SBFM-InvoiceEstmPriceQuote]
WHERE ([__OrderingID] = ID);

PARAMETERS __InvoiceNo Value;
SELECT DISTINCTROW *
FROM QInfobyInvoiceNumber AS [INV-SBFM-ViewInvoice]
WHERE ([__InvoiceNo] = InvoiceNo);

PARAMETERS __OrderingID Value;
SELECT DISTINCTROW *
FROM OrderingAttorneyInfo AS [INV-SBFM-ViewInvoice]
WHERE ([__OrderingID] = ID);

SELECT DISTINCTROW *
FROM Customers;

SELECT Employees.ID, Employees.[Last Name]
FROM Employees;

PARAMETERS [__Order ID] Value;
SELECT DISTINCTROW *
FROM [Order Details Extended] AS Orders
WHERE ([__Order ID] = [Order ID]);

SELECT [ID], [Company]
FROM [Shippers Extended]
ORDER BY [Company];

SELECT [Rates].[ID], [Rates].[Code], [Rates].[List Price], [Rates].[ProductName]
FROM Rates
ORDER BY [Code], [ProductName], [List Price];

SELECT DISTINCTROW *
FROM BrandingThemes;

PARAMETERS __OIFID Value;
SELECT DISTINCTROW *
FROM OrderingAttorneyInfo AS PJOrderingInfoForm
WHERE ([__OIFID] = CourtDatesID);

PARAMETERS __Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField Value;
SELECT DISTINCTROW *
FROM ViewJobFormAppearancesQ AS PJOrderingInfoForm
WHERE ([__Forms]!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField = CourtDates.ID);

PARAMETERS __Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField Value;
SELECT DISTINCTROW *
FROM TempTasksDay1 AS PJStatusChecklist
WHERE ([__Forms]!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField = CourtDates.ID);

PARAMETERS __Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField Value;
SELECT DISTINCTROW *
FROM SBFMCaseInfoQ AS PJStatusChecklist
WHERE ([__Forms]!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField = SBFMCourtDates.ID);

PARAMETERS __Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField Value;
SELECT DISTINCTROW *
FROM [TR-Court-Q] AS PJViewJobForm
WHERE ([__Forms]!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField = CourtDatesID);

PARAMETERS __Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField Value;
SELECT DISTINCTROW *
FROM CustomerAddressA0O1 AS PJViewJobForm
WHERE ([__Forms]!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField = CourtDates.ID);

SELECT [ShippingOptions].[ID], [ShippingOptions].[CourtDatesID], [ShippingOptions].[ToName]
FROM ShippingOptions;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [TurnaroundTimes].[ID], [TurnaroundTimes].[Length]
FROM TurnaroundTimes;

SELECT [UnitPrice].[ID], [UnitPrice].[Rate]
FROM UnitPrice;

SELECT Customers.ID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.JobTitle, Customers.EmailAddress, Customers.City, Customers.State
FROM Customers
ORDER BY Customers.[Company];

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT *
FROM UncompletedStatusesQ
WHERE (UncompletedStatusesQ.CourtDatesID = [Statuses].[ID]);

SELECT [CourtDates].[CasesID]
FROM CourtDates;

SELECT [Cases].[ID]
FROM Cases;

SELECT [Customers].[ID]
FROM Customers;

SELECT [TurnaroundTimes].[ID], [TurnaroundTimes].[Length]
FROM TurnaroundTimes;

SELECT [UnitPrice].[ID], [UnitPrice].[Rate]
FROM UnitPrice;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [CourtDates].[ID]
FROM CourtDates;

SELECT [Invoices].[ID]
FROM Invoices;

SELECT [Customers].[ID]
FROM Customers;

SELECT [Customers].[ID]
FROM Customers;

SELECT [TurnaroundTimes].[ID], [TurnaroundTimes].[Length]
FROM TurnaroundTimes;

SELECT *
FROM ShippingOptions
WHERE [ShippingOptions].[CourtDatesID] = 1755;

SELECT DISTINCTROW *
FROM Doctors;

SELECT AGShortcuts.ID
FROM AGShortcuts;

SELECT DISTINCTROW *
FROM CourtDates;

SELECT DISTINCTROW *
FROM ShippingOptions;

SELECT DISTINCTROW *
FROM TempTasksDay1;

SELECT DISTINCTROW *
FROM Customers;

SELECT [INV-Q-CustomerHistory].[Company], [INV-Q-CustomerHistory].[FirstName], [INV-Q-CustomerHistory].[LastName], [INV-Q-CustomerHistory].[EmailAddress], [INV-Q-CustomerHistory].[AudioLength], [INV-Q-CustomerHistory].[ActualQuantity], [INV-Q-CustomerHistory].[FinalPrice], [INV-Q-CustomerHistory].[InvoiceDate], [INV-Q-CustomerHistory].[InvoiceNo], [Customers].[ID]
FROM [INV-Q-CustomerHistory] INNER JOIN Customers ON [INV-Q-CustomerHistory].[CustomersID] =[Customers].[ID];

SELECT [Cases].[Party1], [Cases].[Party2], [Cases].[CaseNumber1], [Cases].[CaseNumber2], [Cases].[Jurisdiction], [Cases].[HearingTitle], [Cases].[Judge], [Cases].[JudgeTitle], [Cases].[CourtDatesID], [CourtDates].[HearingDate], [CourtDates].[HearingStartTime], [CourtDates].[HearingEndTime], [CourtDates].[App1], [CourtDates].[App2], [CourtDates].[App3], [CourtDates].[App4], [CourtDates].[App5], [CourtDates].[App6], [CourtDates].[OrderingID], [CourtDates].[AudioLength], [CourtDates].[TurnaroundTimesCD], [CourtDates].[InvoicesID], [CourtDates].[DateFactored], [CourtDates].[DatePaid], [CourtDates].[ShipDate], [CourtDates].[ShippingID], [CourtDates].[TrackingNumber]
FROM CourtDates INNER JOIN Cases ON [CourtDates].[ID] =[Cases].[CourtDatesID];

SELECT Tasks.ID, Tasks.CourtDatesID, Tasks.[Due Date], Tasks.Priority, Tasks.Category, Tasks.PriorityPoints, Tasks.Title, Tasks.Description, Tasks.TimeLength, Tasks.Completed, DSum([TimeLength])
FROM Tasks
WHERE ((Tasks.Priority Not Like "*Waiting For*") And (Tasks.Completed=False) And (Tasks.Category Like "Production"))
ORDER BY Tasks.PriorityPoints DESC , Tasks.[Due Date], Tasks.Title;

SELECT AGShortcuts.CourtDatesID, AGShortcuts.AG1, AGShortcuts.ag2, AGShortcuts.ag3, AGShortcuts.ag4, AGShortcuts.ag5, AGShortcuts.ag6, AGShortcuts.ag11, AGShortcuts.ag12, AGShortcuts.ag13, AGShortcuts.ag14, AGShortcuts.ag15, AGShortcuts.ag16, AGShortcuts.ag21, AGShortcuts.ag22, AGShortcuts.ag23, AGShortcuts.ag24, AGShortcuts.ag25, AGShortcuts.ag26, AGShortcuts.ag31, AGShortcuts.ag32, AGShortcuts.ag33, AGShortcuts.ag34, AGShortcuts.ag35, AGShortcuts.ag36, AGShortcuts.ag41, AGShortcuts.ag42, AGShortcuts.ag43, AGShortcuts.ag44, AGShortcuts.ag45, AGShortcuts.ag46, AGShortcuts.ag51, AGShortcuts.ag52, AGShortcuts.ag53, AGShortcuts.ag54, AGShortcuts.ag55, AGShortcuts.ag56, AGShortcuts.ag61, AGShortcuts.ag62, AGShortcuts.ag63, AGShortcuts.ag64, AGShortcuts.ag65, AGShortcuts.ag66
FROM AGShortcuts
WHERE (((AGShortcuts.CourtDatesID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT CourtDates.ID, CourtDates.OrderingID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, Customers.ID, Customers.Company, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.FaxNumber, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.Notes
FROM CourtDates INNER JOIN Customers ON (CourtDates.OrderingID like Customers.ID) OR (CourtDates.App1 like Customers.ID) Or (CourtDates.App2 like Customers.ID) Or (CourtDates.App3 like Customers.ID) Or (CourtDates.App4 like Customers.ID) Or (CourtDates.App5 like Customers.ID) Or (CourtDates.App6 like Customers.ID)
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField) And Customers.ID Like CourtDates.OrderingID Or Customers.ID Like CourtDates.App1 Or Customers.ID Like CourtDates.App2 Or Customers.ID Like CourtDates.App3 Or Customers.ID Like CourtDates.App4 Or Customers.ID Like CourtDates.App5 Or Customers.ID Like CourtDates.App6;

SELECT PaymentQueryInvoiceInfo.FinalPrice, PaymentQueryInvoiceInfo.PaymentsID, PaymentQueryInvoiceInfo.pInvoiceNo, PaymentQueryInvoiceInfo.Amount, PaymentQueryInvoiceInfo.RemitDate, PaymentQueryInvoiceInfo.CourtDatesID, PaymentQueryInvoiceInfo.HearingDate, PaymentQueryInvoiceInfo.HearingStartTime, PaymentQueryInvoiceInfo.HearingEndTime, PaymentQueryInvoiceInfo.CasesID, PaymentQueryInvoiceInfo.OrderingID, PaymentQueryInvoiceInfo.AudioLength, PaymentQueryInvoiceInfo.TurnaroundTimesCD, PaymentQueryInvoiceInfo.DueDate, PaymentQueryInvoiceInfo.cInvoiceNo, PaymentQueryInvoiceInfo.InvoiceDate, PaymentQueryInvoiceInfo.PaymentDueDate, PaymentQueryInvoiceInfo.Subtotal, PaymentQueryInvoiceInfo.UnitPrice, Cases.ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1
FROM Cases INNER JOIN PaymentQueryInvoiceInfo ON Cases.[ID] = PaymentQueryInvoiceInfo.[CasesID];

SELECT Sum([PaymentQueryInvoiceInfo].[Amount]) AS PaymentSum
FROM PaymentQueryInvoiceInfo;

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Statuses.ID AS Statuses_ID, Statuses.CourtDatesID, Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN Statuses ON (CourtDates.[ID] = Statuses.[CourtDatesID]) AND (CourtDates.[StatusesID] = Statuses.[ID])
WHERE (((Statuses.ContactsEntered)=Yes) AND ((Statuses.JobEntered)=Yes) AND ((Statuses.CoverPage)=Yes) AND ((Statuses.AutoCorrect)=Yes) AND ((Statuses.Schedule)=Yes) AND ((Statuses.Invoice)=Yes) AND ((Statuses.Transcribe)=Yes) AND ((Statuses.AddRDtoCover)=Yes) AND ((Statuses.FindReplaceRD)=Yes) AND ((Statuses.HyperlinkTranscripts)=Yes) AND ((Statuses.SpellingsEmail)=Yes) AND ((Statuses.AudioProof)=No)) OR (((Statuses.InvoiceCompleted)=No)) OR (((Statuses.NoticeofService)=No)) OR (((Statuses.PackageEnclosedLetter)=No)) OR (((Statuses.CDLabel)=No)) OR (((Statuses.GenerateZIPs)=No)) OR (((Statuses.TranscriptsReady)=No)) OR (((Statuses.InvoicetoFactorEmail)=No)) OR (((Statuses.FileTranscript)=No)) OR (((Statuses.BurnCD)=No)) OR (((Statuses.ShippingXMLs)=No)) OR (((Statuses.GenerateShippingEM)=No)) OR (((Statuses.AddTrackingNumber)=No));

SELECT CommunicationHistory.CourtDatesID, Format([CommunicationHistory].[DateCreated],"mm/dd/yyyy") AS DateCreated, CommunicationHistory.FileHyperlink1
FROM CommunicationHistory
WHERE CommunicationHistory.CourtDatesID Is Null;

SELECT CommunicationHistory.ID, CommunicationHistory.FileHyperlink, CommunicationHistory.FileHyperlink1, CommunicationHistory.DateCreated, CommunicationHistory.CourtDatesID, CommunicationHistory.CustomersID
FROM CommunicationHistory
ORDER BY CommunicationHistory.CourtDatesID;

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.OrderingID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.EmailAddress, Customers.Address, Customers.City, Customers.State, Customers.ZIP, ShippingOptions.ID AS ShippingOptions_ID, ShippingOptions.CourtDatesID, ShippingOptions.CourtDatesIDLK, ShippingOptions.MailClass, ShippingOptions.PackageType, ShippingOptions.Width, ShippingOptions.Length, ShippingOptions.Depth, ShippingOptions.PriorityMailExpress1030, ShippingOptions.HolidayDelivery, ShippingOptions.SundayDelivery, ShippingOptions.SaturdayDelivery, ShippingOptions.SignatureRequired, ShippingOptions.Stealth, ShippingOptions.ReplyPostage, ShippingOptions.InsuredMail, ShippingOptions.COD, ShippingOptions.RestrictedDelivery, ShippingOptions.AdultSignatureRestricted, ShippingOptions.AdultSignatureRequired, ShippingOptions.ReturnReceipt, ShippingOptions.CertifiedMail, ShippingOptions.SignatureConfirmation, ShippingOptions.USPSTracking, ShippingOptions.ReferenceID, ShippingOptions.ToName, ShippingOptions.ToAddress1, ShippingOptions.ToAddress2, ShippingOptions.ToCity, ShippingOptions.ToState, ShippingOptions.ToPostalCode, ShippingOptions.ToCountry, ShippingOptions.Value, ShippingOptions.Description, ShippingOptions.ToEMail, ShippingOptions.ToPhone, ShippingOptions.WeightOz, ShippingOptions.ActualWeight, ShippingOptions.ActualWeightText, ShippingOptions.Amount
FROM (Customers INNER JOIN CourtDates ON Customers.[ID] = CourtDates.[OrderingID]) INNER JOIN ShippingOptions ON CourtDates.[ID] = ShippingOptions.[CourtDatesIDLK];

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.ShippingID, Customers.ID AS Customers_ID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.BusinessPhone, Customers.EmailAddress, ShippingOptions.ID AS ShippingOptions_ID, ShippingOptions.CourtDatesID, ShippingOptions.MailClass, ShippingOptions.PackageType, ShippingOptions.Width, ShippingOptions.Length, ShippingOptions.Depth, ShippingOptions.PriorityMailExpress1030, ShippingOptions.HolidayDelivery, ShippingOptions.SundayDelivery, ShippingOptions.SaturdayDelivery, ShippingOptions.SignatureRequired, ShippingOptions.Stealth, ShippingOptions.ReplyPostage, ShippingOptions.InsuredMail, ShippingOptions.COD, ShippingOptions.RestrictedDelivery, ShippingOptions.AdultSignatureRestricted, ShippingOptions.AdultSignatureRequired, ShippingOptions.ReturnReceipt, ShippingOptions.CertifiedMail, ShippingOptions.SignatureConfirmation, ShippingOptions.USPSTracking, ShippingOptions.ReferenceID, ShippingOptions.ToName, ShippingOptions.ToAddress1, ShippingOptions.ToAddress2, ShippingOptions.ToCity, ShippingOptions.ToState, ShippingOptions.ToPostalCode, ShippingOptions.ToCountry, ShippingOptions.Value, ShippingOptions.Description, ShippingOptions.ToEMail, ShippingOptions.ToPhone, ShippingOptions.WeightOz, ShippingOptions.ActualWeight, ShippingOptions.ActualWeightText, ShippingOptions.Amount
FROM (Customers INNER JOIN CourtDates ON Customers.[ID] = CourtDates.[OrderingID]) INNER JOIN ShippingOptions ON CourtDates.[ID] = ShippingOptions.[CourtDatesID]
WHERE ShippingOptions.[CourtDatesID]=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField];

SELECT CourtDates.ID AS CourtDatesID, CourtDates.BrandingTheme AS CourtDates_BrandingTheme, BrandingThemes.ID AS BrandingThemes_ID, BrandingThemes.BrandingTheme AS BrandingThemes_BrandingTheme
FROM BrandingThemes INNER JOIN CourtDates ON BrandingThemes.[ID] = CourtDates.[BrandingTheme];

SELECT CourtDatesBTRQuery2.BrandingThemes_BrandingTheme, InvoicesQuery4.CourtDatesID AS InvoicesQuery4_CourtDatesID, InvoicesQuery4.Reference, InvoicesQuery4.HearingDate, InvoicesQuery4.HearingStartTime, InvoicesQuery4.HearingEndTime, InvoicesQuery4.CasesID, InvoicesQuery4.OrderingID, InvoicesQuery4.AudioLength, InvoicesQuery4.Location, InvoicesQuery4.TurnaroundTimesCD, InvoicesQuery4.Expr1010, InvoicesQuery4.Cases_ID, InvoicesQuery4.Party1, InvoicesQuery4.Party2, InvoicesQuery4.CaseNumber1, InvoicesQuery4.CaseNumber2, InvoicesQuery4.Jurisdiction, InvoicesQuery4.CustomersID, InvoicesQuery4.Company, InvoicesQuery4.FirstName, InvoicesQuery4.LastName, InvoicesQuery4.Address, InvoicesQuery4.City, InvoicesQuery4.State, InvoicesQuery4.ZIP, InvoicesQuery4.EmailAddress, InvoicesQuery4.InvoiceNo, InvoicesQuery4.Quantity, InvoicesQuery4.InventoryItemCode, InvoicesQuery4.DueDate, InvoicesQuery4.InvoiceDate, InvoicesQuery4.AccountCode, InvoicesQuery4.TaxType, InvoicesQuery4.BrandingTheme, CourtDatesBTRQuery2.CourtDatesID, CourtDatesBTRQuery2.Code, CourtDatesBTRQuery2.[Rate]
FROM InvoicesQuery4 INNER JOIN CourtDatesBTRQuery2 ON InvoicesQuery4.CourtDatesID=CourtDatesBTRQuery2.[CourtDatesID];

SELECT CourtDatesBTQuery.CourtDates_ID, CourtDatesBTQuery.BrandingThemes_BrandingTheme, CourtDatesRatesQuery.CourtDatesID AS CourtDatesRatesQuery_CourtDatesID, InvoicesQuery4.CourtDatesID AS InvoicesQuery4_CourtDatesID, InvoicesQuery4.Reference, InvoicesQuery4.HearingDate, InvoicesQuery4.HearingStartTime, InvoicesQuery4.HearingEndTime, InvoicesQuery4.CasesID, InvoicesQuery4.OrderingID, InvoicesQuery4.AudioLength, InvoicesQuery4.Location, InvoicesQuery4.TurnaroundTimesCD, InvoicesQuery4.Expr1010, InvoicesQuery4.Cases_ID, InvoicesQuery4.Party1, InvoicesQuery4.Party2, InvoicesQuery4.CaseNumber1, InvoicesQuery4.CaseNumber2, InvoicesQuery4.Jurisdiction, InvoicesQuery4.CustomersID, InvoicesQuery4.Company, InvoicesQuery4.FirstName, InvoicesQuery4.LastName, InvoicesQuery4.Address, InvoicesQuery4.City, InvoicesQuery4.State, InvoicesQuery4.ZIP, InvoicesQuery4.EmailAddress, InvoicesQuery4.InvoiceNo, InvoicesQuery4.Quantity, InvoicesQuery4.DueDate, InvoicesQuery4.InvoiceDate, InvoicesQuery4.AccountCode, InvoicesQuery4.TaxType, CourtDatesRatesQuery.Code, CourtDatesRatesQuery.[List Price]
FROM CourtDatesRatesQuery INNER JOIN (CourtDatesBTQuery INNER JOIN InvoicesQuery4 ON CourtDatesBTQuery.[CourtDatesID] = InvoicesQuery4.[CourtDatesID]) ON CourtDatesRatesQuery.[CourtDatesID] = InvoicesQuery4.[CourtDatesID];

SELECT CourtDatesBTQuery.CourtDatesID, CourtDatesBTQuery.BrandingThemes_BrandingTheme, CourtDatesRatesQuery.CourtDatesID AS CourtDatesRatesQuery_CourtDatesID, InvoicesQuery4.CourtDatesID AS InvoicesQuery4_CourtDatesID, InvoicesQuery4.Reference, InvoicesQuery4.HearingDate, InvoicesQuery4.HearingStartTime, InvoicesQuery4.HearingEndTime, InvoicesQuery4.CasesID, InvoicesQuery4.OrderingID, InvoicesQuery4.AudioLength, InvoicesQuery4.Location, InvoicesQuery4.TurnaroundTimesCD, InvoicesQuery4.Expr1010, InvoicesQuery4.Cases_ID, InvoicesQuery4.Party1, InvoicesQuery4.Party2, InvoicesQuery4.CaseNumber1, InvoicesQuery4.CaseNumber2, InvoicesQuery4.Jurisdiction, InvoicesQuery4.CustomersID, InvoicesQuery4.Company, InvoicesQuery4.FirstName, InvoicesQuery4.LastName, InvoicesQuery4.Address, InvoicesQuery4.City, InvoicesQuery4.State, InvoicesQuery4.ZIP, InvoicesQuery4.EmailAddress, InvoicesQuery4.InvoiceNo, InvoicesQuery4.Quantity, InvoicesQuery4.DueDate, InvoicesQuery4.InvoiceDate, InvoicesQuery4.AccountCode, InvoicesQuery4.TaxType, CourtDatesRatesQuery.Code, CourtDatesRatesQuery.[Rate]
FROM CourtDatesRatesQuery INNER JOIN (CourtDatesBTQuery INNER JOIN InvoicesQuery4 ON CourtDatesBTQuery.[CourtDatesID] = InvoicesQuery4.[CourtDatesID]) ON CourtDatesRatesQuery.[CourtDatesID] = InvoicesQuery4.[CourtDatesID];

SELECT CourtDates.ID, CourtDates.HearingDate, CourtDates.HearingStartTime, [CourtDates].HearingEndTime
FROM CourtDates
WHERE (CourtDates.[ID])=forms![MMProcess Jobs].JobNumberField
ORDER BY [HearingDate], [HearingStartTime], [HearingEndTime];

SELECT CourtDates.ID AS CourtDatesID, CourtDates.InventoryRateCode, Rates.ID AS RatesID, Rates.Code, Rates.[List Price] AS Rate
FROM CourtDates INNER JOIN Rates ON CourtDates.[InventoryRateCode]=Rates.[ID];

SELECT CourtDates.InvoiceNo, "0" AS TotalExpenses
FROM CourtDates LEFT JOIN Expenses ON CourtDates.[InvoiceNo] = Expenses.[InvoiceNo]
WHERE (((Expenses.InvoiceNo) Is Null));

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.OrderingID
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App1
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App2
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App3
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App4
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App5
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT *
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.App6
WHERE ((CourtDates.ID) Like Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField);

SELECT ViewJobFormAppearancesQ.MrMs, ViewJobFormAppearancesQ.LastName
FROM ViewJobFormAppearancesQ
WHERE (((ViewJobFormAppearancesQ.CourtDates.ID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT CommunicationHistory.CourtDatesID, CommunicationHistory.FileHyperlink, CommunicationHistory.DateCreated
FROM CommunicationHistory
WHERE (((CommunicationHistory.CourtDatesID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT *
FROM CommunicationHistory;

SELECT QTotalExpensesByInvoiceReal.[TotalExpenses], QTotalPricebyInvoiceNumber.[TotalPrice], QTotalPageCountbyInvoiceNumber.[PageCount], QTotalPricebyInvoiceNumber.[InvoiceNo], QTotalExpensesByInvoiceReal.[InvoiceNo], QTotalPageCountbyInvoiceNumber.[InvoiceNo], QTotalInvoiceNumberDate.[InvoiceNo], QTotalInvoiceNumberDate.[InvoiceDate]
FROM ((QTotalPageCountbyInvoiceNumber INNER JOIN QTotalExpensesByInvoiceReal ON QTotalPageCountbyInvoiceNumber.InvoiceNo = QTotalExpensesByInvoiceReal.InvoiceNo) INNER JOIN QTotalPricebyInvoiceNumber ON QTotalPageCountbyInvoiceNumber.InvoiceNo = QTotalPricebyInvoiceNumber.InvoiceNo) INNER JOIN QTotalInvoiceNumberDate ON QTotalPageCountbyInvoiceNumber.InvoiceNo = QTotalInvoiceNumberDate.InvoiceNo;

SELECT ContactName, InvoiceNumber, Reference AS [PO Number], InvoiceDate AS [Invoice Date], "28" AS [Net Term], ActualQuantity * UnitAmount AS [Invoice Amount]
FROM XeroInvoiceCSV
WHERE Reference=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField];

SELECT FinalUnitPriceQuery.CourtDatesID, FinalUnitPriceQuery.ID AS FinalUnitPriceQueryID, FinalUnitPriceQuery.FinalPrice AS FinalPrice, FinalUnitPriceQuery.AudioLength AS FinalUnitPriceQuery_AudioLength, FinalUnitPriceQuery.TurnaroundTimesCD AS FinalUnitPriceQuery_TurnaroundTimesCD, FinalUnitPriceQuery.InvoiceNo AS FinalUnitPriceQuery_InvoiceNo, FinalUnitPriceQuery.InvoiceDate AS FinalUnitPriceQuery_InvoiceDate, FinalUnitPriceQuery.PaymentDueDate AS FinalUnitPriceQuery_PaymentDueDate, FinalUnitPriceQuery.ExpectedAdvanceDate AS FinalUnitPriceQuery_ExpectedAdvanceDate, FinalUnitPriceQuery.ExpectedRebateDate AS FinalUnitPriceQuery_ExpectedRebateDate, FinalUnitPriceQuery.EstimatedPageCount AS FinalUnitPriceQuery_EstimatedPageCount, FinalUnitPriceQuery.UnitPrice AS FinalUnitPriceQuery_UnitPrice, FinalUnitPriceQuery.Rate, FinalUnitPriceQuery.ActualQuantity AS FinalUnitPriceQuery_ActualQuantity, FinalUnitPriceQuery.DueDate AS FinalUnitPriceQuery_DueDate, InvoiceInfoQ.ID AS InvoiceInfoQ_ID, InvoiceInfoQ.FactoringApproved, InvoiceInfoQ.MrMs, InvoiceInfoQ.Company, InvoiceInfoQ.LastName, InvoiceInfoQ.FirstName, InvoiceInfoQ.BusinessPhone, InvoiceInfoQ.EmailAddress, InvoiceInfoQ.Address, InvoiceInfoQ.City, InvoiceInfoQ.State, InvoiceInfoQ.ZIP, InvoiceInfoQ.HearingDate, InvoiceInfoQ.HearingStartTime, InvoiceInfoQ.HearingEndTime, InvoiceInfoQ.CasesID, InvoiceInfoQ.OrderingID, InvoiceInfoQ.AudioLength AS InvoiceInfoQ_AudioLength, InvoiceInfoQ.Location, InvoiceInfoQ.TurnaroundTimesCD AS InvoiceInfoQ_TurnaroundTimesCD, InvoiceInfoQ.Subtotal, InvoiceInfoQ.DueDate AS InvoiceInfoQ_DueDate, InvoiceInfoQ.InvoiceNo AS InvoiceNo, InvoiceInfoQ.FiledNotFiled, InvoiceInfoQ.InvoiceDate AS InvoiceInfoQ_InvoiceDate, InvoiceInfoQ.PaymentDueDate AS InvoiceInfoQ_PaymentDueDate, InvoiceInfoQ.Quantity, InvoiceInfoQ.ActualQuantity AS InvoiceInfoQ_ActualQuantity, InvoiceInfoQ.UnitPrice AS InvoiceInfoQ_UnitPrice, InvoiceInfoQ.ExpectedAdvanceDate AS InvoiceInfoQ_ExpectedAdvanceDate, InvoiceInfoQ.ExpectedRebateDate AS InvoiceInfoQ_ExpectedRebateDate, InvoiceInfoQ.EstimatedPageCount AS InvoiceInfoQ_EstimatedPageCount, InvoiceInfoQ.Party1, InvoiceInfoQ.Party2, InvoiceInfoQ.CaseNumber1, InvoiceInfoQ.CaseNumber2, InvoiceInfoQ.Jurisdiction, InvoiceInfoQ.HearingTitle, InvoiceInfoQ.Judge, InvoiceInfoQ.JudgeTitle, InvoiceInfoQ.ShipDate, InvoiceInfoQ.TrackingNumber, InvoiceInfoQ.FactoringCost AS FactoringCost, InvoiceInfoQ.Factored, InvoiceInfoQ.FinalPrice AS FinalPrice1
FROM FinalUnitPriceQuery INNER JOIN InvoiceInfoQ ON FinalUnitPriceQuery.[InvoiceNo] = InvoiceInfoQ.[InvoiceNo]
WHERE FinalUnitPriceQuery.[CourtDatesID] = InvoiceInfoQ.[CourtDatesID];

SELECT CourtDates.ID AS CourtDatesID, UnitPrice.ID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.ExpectedAdvanceDate, CourtDates.ExpectedRebateDate, CourtDates.EstimatedPageCount, CourtDates.FactoringCost, CourtDates.UnitPrice, UnitPrice.Rate, CourtDates.ActualQuantity, CourtDates.DueDate, CourtDates.FinalPrice AS FinalPrice, Rate*ActualQuantity AS Subtotal
FROM CourtDates INNER JOIN UnitPrice ON CourtDates.[UnitPrice] = UnitPrice.[ID];

SELECT FinalUnitPriceQuery.InvoiceNo, FinalUnitPriceQuery.InvoiceDate, FinalUnitPriceQuery.Subtotal
FROM FinalUnitPriceQuery;

SELECT Cases.[Party1], Cases.[Party2], Cases.[CaseNumber1], Cases.[CaseNumber2]
FROM Cases
WHERE (((Cases.[Party1]) In (SELECT [Party1] FROM [Cases] As Tmp GROUP BY [Party1],[Party2] HAVING Count(*)>1  And [Party2] = [Cases].[Party2])))
ORDER BY Cases.[Party1], Cases.[Party2];

SELECT First(Payments.[InvoiceNo]) AS [InvoiceNo Field], First(Payments.[Amount]) AS [Amount Field], First(Payments.[RemitDate]) AS [RemitDate Field], Count(Payments.[InvoiceNo]) AS NumberOfDups
FROM Payments
GROUP BY Payments.[InvoiceNo], Payments.[Amount], Payments.[RemitDate]
HAVING (((Count(Payments.[InvoiceNo]))>1) AND ((Count(Payments.[RemitDate]))>1));

SELECT *
FROM CourtDates INNER JOIN OrderingAttorneyInfo ON CourtDates.ID=OrderingAttorneyInfo.CourtDatesID;

SELECT Tasks.ID, Tasks.CourtDatesID, Tasks.[Due Date], Tasks.Priority, Tasks.Category, Tasks.PriorityPoints, Tasks.Title, Tasks.Description, Tasks.TimeLength, Tasks.Completed
FROM Tasks
WHERE ((Tasks.Priority Not Like "*Waiting For*") And (Tasks.Completed=False))
ORDER BY Tasks.PriorityPoints DESC , Tasks.[Due Date], Tasks.Title;

TRANSFORM Sum(GroupTasksIncomplete.TimeLength) AS SumOfTimeLength
SELECT GroupTasksIncomplete.PriorityPoints, Sum(GroupTasksIncomplete.TimeLength) AS [Total Of TimeLength]
FROM GroupTasksIncomplete
GROUP BY GroupTasksIncomplete.PriorityPoints
PIVOT GroupTasksIncomplete.Completed;

SELECT Tasks.ID, Tasks.CourtDatesID, Tasks.[Due Date], Tasks.Priority, Tasks.Category, Tasks.PriorityPoints, Tasks.Title, Tasks.Description, Tasks.TimeLength, Tasks.Completed
FROM Tasks
WHERE ((Tasks.Priority Not Like "*Waiting For*") And (Tasks.Completed=False) And (Tasks.Category Like "Production"))
ORDER BY Tasks.PriorityPoints DESC , Tasks.[Due Date], Tasks.Title;

SELECT GroupTasksIncompleteProduction.PriorityPoints, GroupTasksIncompleteProduction.TimeLength
FROM GroupTasksIncompleteProduction;

SELECT CourtDates.ID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.DatePaid, CourtDates.DueDate, Cases.Jurisdiction, UncompletedStatusesQ.CourtDatesID
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN UncompletedStatusesQ ON CourtDates.[ID] = UncompletedStatusesQ.[CourtDatesID];

SELECT CourtDates.InvoiceDate, CourtDates.ID AS CourtDatesID, (DateAdd('d',28,InvoiceDate)) AS ["PaymentDueDate"]
FROM CourtDates
WHERE CourtDates.ID =Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField];

SELECT Customers.ID, Customers.FactoringApproved, Customers.MrMs, Customers.[Company], Customers.[LastName], Customers.[FirstName], Customers.[BusinessPhone], Customers.EmailAddress, Customers.Address, Customers.City, Customers.[State], Customers.[ZIP], Courtdates.HearingDate, Courtdates.HearingStartTime, Courtdates.HearingEndTime, Courtdates.CasesID, Courtdates.OrderingID, Courtdates.AudioLength, Courtdates.Location, Courtdates.TurnaroundTimesCD, CourtDates.Subtotal, CourtDates.Quantity, Courtdates.DueDate, Courtdates.InvoiceNo, CourtDates.ID AS CourtDatesID, Courtdates.FiledNotFiled, Courtdates.InvoiceDate, Courtdates.PaymentDueDate, Courtdates.Quantity, CourtDates.ActualQuantity, CourtDates.UnitPrice, Courtdates.ExpectedAdvanceDate, Courtdates.ExpectedRebateDate, Courtdates.EstimatedPageCount, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, CourtDates.ShipDate, CourtDates.FactoringCost, CourtDates.TrackingNumber, CourtDates.Factored, CourtDates.FinalPrice
FROM Customers INNER JOIN (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) ON Customers.[ID] = CourtDates.[OrderingID];

SELECT CourtDates.InvoiceDate AS InvoiceDate, CourtDates.ID AS CourtDatesID, (DateAdd('d',1,InvoiceDate)) AS PaymentDueDate
FROM CourtDates INNER JOIN UnitPrice ON CourtDates.[UnitPrice] = UnitPrice.[ID];

SELECT CourtDates.ID, CourtDates.HearingDate, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.Location, Orders.InventoryItemCode, Cases.ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.Judge, Cases.HearingTitle, Rates.Code, Rates.ProductName, Rates.[List Price], Orders.OrderDate, Orders.DateShip, Orders.DateFactored, Orders.PaymentType, Orders.DatePaid, Orders.CourtDatesID, Orders.InvoiceNumber, Orders.Quantity, Orders.Reference, Orders.DueDate, Orders.InvoiceDate, Orders.InventoryItemCode, Orders.AccountCode, Orders.TaxType, Orders.BrandingTheme, Customers.ID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.EmailAddress, Customers.Address, Customers.City, Customers.State, Customers.ZIP
FROM Customers INNER JOIN ((Cases INNER JOIN CourtDates ON CourtDates.[CasesID] = Cases.[ID]) INNER JOIN (Rates INNER JOIN Orders ON Rates.[Code] = Orders.[InventoryItemCode]) ON CourtDates.[ID] = Orders.[CourtDatesID]) ON Customers.[ID] = CourtDates.[OrderingID]
WHERE (((CourtDates.ID)=Orders.Reference) And ((CourtDates.CasesID)=Cases.ID) And ((CourtDates.OrderingID)=Customers.ID) And ((Orders.Reference)=Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField));

SELECT OrderingAttorneyInfo.ID AS OrderingAttorneyInfo_ID, OrderingAttorneyInfo.Company, OrderingAttorneyInfo.MrMs, OrderingAttorneyInfo.LastName, OrderingAttorneyInfo.FirstName, OrderingAttorneyInfo.EmailAddress, OrderingAttorneyInfo.BusinessPhone, OrderingAttorneyInfo.Address, OrderingAttorneyInfo.City, OrderingAttorneyInfo.State, OrderingAttorneyInfo.ZIP, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, CourtDates.HearingDate, CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.InvoicesID, Rates.ID AS Rates_ID, Rates.Code, Rates.ProductName, Rates.[List Price], Orders.[Order ID], Orders.OrderDate, Orders.DateShip, Orders.DateFactored, Orders.PaymentType, Orders.DatePaid, Orders.CourtDatesID, Orders.InvoiceNumber, Orders.Quantity, Orders.InventoryItemCode, Orders.Reference, Orders.DueDate, Orders.InvoiceDate, Orders.AccountCode, Orders.TaxType, Orders.BrandingTheme
FROM ((Cases INNER JOIN OrderingAttorneyInfo ON Cases.[ID] = OrderingAttorneyInfo.[CasesID]) INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN (Rates INNER JOIN Orders ON Rates.[ID] = Orders.[InventoryItemCode]) ON CourtDates.[ID] = Orders.[CourtDatesID];

SELECT OrderingAttorneyInfo.ID AS OrderingAttorneyInfo_ID, OrderingAttorneyInfo.Company, OrderingAttorneyInfo.MrMs, OrderingAttorneyInfo.LastName, OrderingAttorneyInfo.FirstName, OrderingAttorneyInfo.EmailAddress, OrderingAttorneyInfo.BusinessPhone, OrderingAttorneyInfo.Address, OrderingAttorneyInfo.City, OrderingAttorneyInfo.State, OrderingAttorneyInfo.ZIP, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, CourtDates.HearingDate, CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.InvoicesID, Rates.ID AS Rates_ID, Rates.Code, Rates.ProductName, Rates.[List Price], Orders.[Order ID], Orders.OrderDate, Orders.DateShip, Orders.DateFactored, Orders.PaymentType, Orders.DatePaid, Orders.CourtDatesID, Orders.InvoiceNumber, Orders.Quantity, Orders.InventoryItemCode, Orders.Reference, Orders.DueDate, Orders.InvoiceDate, Orders.AccountCode, Orders.TaxType, Orders.BrandingTheme AS BrandingTheme
FROM ((Cases INNER JOIN OrderingAttorneyInfo ON Cases.[ID] = OrderingAttorneyInfo.[CasesID]) INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN (Rates INNER JOIN Orders ON Rates.[ID] = Orders.[InventoryItemCode]) ON CourtDates.[ID] = Orders.[CourtDatesID];

SELECT CourtDates.ID AS CourtDatesID, CourtDates.ID AS Reference, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.Location, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Customers.ID AS CustomersID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.EmailAddress, CourtDates.InvoiceNo, CourtDates.Quantity, CourtDates.InventoryRateCode AS InventoryItemCode, CourtDates.DueDate, CourtDates.InvoiceDate, CourtDates.AccountCode, CourtDates.TaxType, CourtDates.BrandingTheme, CourtDates.FinalPrice, CourtDates.ActualQuantity
FROM Customers INNER JOIN (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) ON Customers.[ID] = CourtDates.[OrderingID];

SELECT InvoicesQuery4.CourtDatesID, InvoicesQuery4.BrandingTheme AS IQ4BrandingTheme, BrandingThemes.ID, BrandingThemes.BrandingTheme AS BTBrandingTheme
FROM InvoicesQuery4 INNER JOIN BrandingThemes ON InvoicesQuery4.[BrandingTheme]=BrandingThemes.[ID];

SELECT CourtDates.ID AS CourtDatesID, CourtDates.ID AS Reference, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.Location, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Customers.ID AS CustomersID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.EmailAddress, CourtDates.InvoiceNo, CourtDates.Quantity, CourtDates.InventoryRateCode AS InventoryItemCode, CourtDates.DueDate, CourtDates.InvoiceDate, CourtDates.AccountCode, CourtDates.TaxType, CourtDates.BrandingTheme, CourtDates.subtotal, CourtDates.FinalPrice, CourtDates.ActualQuantity
FROM Customers INNER JOIN (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) ON Customers.[ID] = CourtDates.[OrderingID]
WHERE CourtDates.FinalPrice = 0;

SELECT Customers.ID AS CustomersID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.BusinessPhone, Customers.EmailAddress, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.FactorApvlID, Customers.FactoringApproved, CourtDates.ID AS CourtDates_ID, CourtDates.HearingDate, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, Cases.Party1, Cases.ID AS Cases_ID, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, CourtDates.FinalPrice, CourtDates.ActualQuantity, CourtDates.InvoiceDate, Payments.Amount, Payments.RemitDate, CourtDates.InvoiceNo
FROM ((CourtDates INNER JOIN Cases ON CourtDates.CasesID=Cases.ID) INNER JOIN Customers ON CourtDates.[OrderingID] =[Customers].[ID]) INNER JOIN Payments ON CourtDates.InvoiceNo=Payments.InvoiceNo;

SELECT CourtDates.ID AS CourtDatesID, CourtDates.InvoiceDate, CourtDates.ActualQuantity, CourtDates.FinalPrice, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceNo, CourtDates.CasesID, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Customers.ID AS Customers_ID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.FactoringApproved, Expenses.ID AS ExpensesID, Expenses.Vendor, Expenses.ExpensesDate, Expenses.Amount, Expenses.Memo, Expenses.CourtDatesID, Expenses.InvoiceNo
FROM Customers INNER JOIN (Cases INNER JOIN (Expenses INNER JOIN CourtDates ON Expenses.[InvoiceNo] = CourtDates.[InvoiceNo]) ON Cases.[ID] = CourtDates.[CasesID]) ON Customers.[ID] = CourtDates.[OrderingID];

UPDATE CourtDates INNER JOIN FinalUnitPriceQuery ON CourtDates.ID = FinalUnitPriceQuery.CourtDatesID SET CourtDates.FinalPrice = ([FinalUnitPriceQuery].[ActualQuantity]*[FinalUnitPriceQuery].[Rate])
WHERE (((CourtDates.ID)=[FinalUnitPriceQuery].[CourtDatesID]));

UPDATE TempCourtDates INNER JOIN TUnitPriceQuery ON TempCourtDates.CourtDatesID = TUnitPriceQuery.CourtDatesID SET TempCourtDates.Subtotal = TUnitPriceQuery.[Subtotal], TempCourtDates.ExpectedAdvanceDate = TUnitPriceQuery.[ExpectedAdvanceDate], TempCourtDates.ExpectedRebateDate = TUnitPriceQuery.[ExpectedRebateDate]
WHERE (([TempCourtDates].[CourtDatesID]=[TUnitPriceQuery].[CourtDatesID]));

UPDATE CourtDates INNER JOIN UnitPriceQuery ON CourtDates.ID = UnitPriceQuery.CourtDatesID SET CourtDates.Subtotal = UnitPriceQuery.[Subtotal], CourtDates.ExpectedAdvanceDate = UnitPriceQuery.[ExpectedAdvanceDate], CourtDates.ExpectedRebateDate = UnitPriceQuery.[ExpectedRebateDate]
WHERE (([CourtDates].[ID]=([UnitPriceQuery].[CourtDatesID])));

SELECT Rates.ID AS Rates_ID, Rates.Code, Rates.[List Price], XeroInvoiceCSV.ID AS XeroInvoiceCSV_ID, XeroInvoiceCSV.InventoryItemCode, XeroInvoiceCSV.UnitAmount
FROM Rates INNER JOIN XeroInvoiceCSV ON Rates.[List Price] = XeroInvoiceCSV.[UnitAmount];

SELECT Rates.ID, Rates.Code, Rates.[List Price], CourtDates.InventoryItemCode, CourtDates.ID
FROM Rates INNER JOIN CourtDates ON Rates.[ID] = CourtDates.[InventoryItemCode];

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Statuses.ID AS Statuses_ID, Statuses.CourtDatesID, Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN Statuses ON (CourtDates.[ID] = Statuses.[CourtDatesID]) AND (CourtDates.[StatusesID] = Statuses.[ID])
WHERE ((Statuses.ContactsEntered)=No) Or ((Statuses.JobEntered)=No) Or ((Statuses.CoverPage)=No) Or ((Statuses.AutoCorrect)=No) Or ((Statuses.Schedule)=No) Or ((Statuses.Invoice)=No);

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Statuses.ID AS Statuses_ID, Statuses.CourtDatesID, Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN Statuses ON (CourtDates.[StatusesID] = Statuses.[ID]) AND (CourtDates.[ID] = Statuses.[CourtDatesID])
WHERE (((Statuses.ContactsEntered)=Yes) AND ((Statuses.JobEntered)=Yes) AND ((Statuses.CoverPage)=Yes) AND ((Statuses.AutoCorrect)=Yes) AND ((Statuses.Schedule)=Yes) AND ((Statuses.Invoice)=Yes) AND ((Statuses.Transcribe)=No)) OR (((Statuses.AddRDtoCover)=No)) OR (((Statuses.FindReplaceRD)=No)) OR (((Statuses.HyperlinkTranscripts)=No)) OR (((Statuses.SpellingsEmail)=No));

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Statuses.ID AS Statuses_ID, Statuses.CourtDatesID, Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN Statuses ON CourtDates.[ID] = Statuses.[CourtDatesID]
WHERE (((Statuses.ContactsEntered)=Yes) AND ((Statuses.JobEntered)=Yes) AND ((Statuses.CoverPage)=Yes) AND ((Statuses.AutoCorrect)=Yes) AND ((Statuses.Schedule)=Yes) AND ((Statuses.Invoice)=Yes) AND ((Statuses.Transcribe)=Yes) AND ((Statuses.AudioProof)=No) AND ((Statuses.InvoiceCompleted)=No));

SELECT CourtDates.ID AS CourtDates_ID, CourtDates.CasesID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, Cases.ID AS Cases_ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Statuses.ID AS Statuses_ID, Statuses.CourtDatesID, Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber
FROM (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) INNER JOIN Statuses ON (CourtDates.[StatusesID] = Statuses.[ID]) AND (CourtDates.[ID] = Statuses.[CourtDatesID])
WHERE (((Statuses.ContactsEntered)=Yes) AND ((Statuses.JobEntered)=Yes) AND ((Statuses.CoverPage)=Yes) AND ((Statuses.AutoCorrect)=Yes) AND ((Statuses.Schedule)=Yes) AND ((Statuses.Invoice)=Yes) AND ((Statuses.Transcribe)=Yes) AND ((Statuses.AddRDtoCover)=Yes) AND ((Statuses.FindReplaceRD)=Yes) AND ((Statuses.HyperlinkTranscripts)=Yes) AND ((Statuses.SpellingsEmail)=Yes) AND ((Statuses.AudioProof)=Yes) AND ((Statuses.InvoiceCompleted)=No) OR ((Statuses.AddTrackingNumber)=No));

SELECT MaxCourtDatesUnion.ID AS MaxCourtDatesUnion_ID, MaxCourtDatesUnion.HearingDate, MaxCourtDatesUnion.HearingStartTime, MaxCourtDatesUnion.HearingEndTime, MaxCourtDatesUnion.CasesID, MaxCourtDatesUnion.App1, MaxCourtDatesUnion.App2, MaxCourtDatesUnion.App3, MaxCourtDatesUnion.App4, MaxCourtDatesUnion.App5, MaxCourtDatesUnion.App6, MaxCourtDatesUnion.OrderingID, MaxCourtDatesUnion.StatusesID, MaxCourtDatesUnion.AudioLength, MaxCourtDatesUnion.Location, MaxCourtDatesUnion.TurnaroundTimesCD, MaxCourtDatesUnion.InvoiceNo, MaxCourtDatesUnion.DueDate, MaxCourtDatesUnion.ShipDate, MaxCourtDatesUnion.TrackingNumber, MaxCourtDatesUnion.PaymentType, MaxCourtDatesUnion.Notes, MaxCourtDatesUnion.ShippingOptionsID, MaxCourtDatesUnion.SPKRID, MaxCourtDatesUnion.AGShortcuts, MaxCourtDatesUnion.CDLabel, MaxCourtDatesUnion.RoughDraft, MaxCourtDatesUnion.AutoCorrect, MaxCourtDatesUnion.PackageEnclosedLetter, MaxCourtDatesUnion.TranscriptsReady, MaxCourtDatesUnion.FactorCustApvl, MaxCourtDatesUnion.CoverPage, MaxCourtDatesUnion.FactorInvFactor, MaxCourtDatesUnion.FiledNotFiled, MaxCourtDatesUnion.Factored, MaxCourtDatesUnion.InvoiceDate, MaxCourtDatesUnion.PaymentDueDate, MaxCourtDatesUnion.FactoringInterestID, MaxCourtDatesUnion.ExpectedRebateDate, MaxCourtDatesUnion.EstimatedPageCount, MaxCourtDatesUnion.FactoringCost, MaxCourtDatesUnion.UnitPrice, MaxCourtDatesUnion.Quantity, MaxCourtDatesUnion.ActualQuantity, MaxCourtDatesUnion.Subtotal, MaxCourtDatesUnion.ExpectedAdvanceDate, MaxCourtDatesUnion.FinalPrice, MaxCourtDatesUnion.PaymentSum, MaxCourtDatesUnion.InventoryRateCode, MaxCourtDatesUnion.AccountCode, MaxCourtDatesUnion.TaxType, MaxCourtDatesUnion.BrandingTheme, MaxCourtDatesUnion.CourtDatesID, Cases.ID AS CasesID, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Customers.ID AS Customers_ID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.JobTitle, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.FactoringApproved
FROM Customers INNER JOIN (Cases INNER JOIN MaxCourtDatesUnion ON Cases.[ID] = MaxCourtDatesUnion.[CasesID]) ON Customers.[ID] = MaxCourtDatesUnion.[OrderingID];

SELECT *
FROM CourtDates INNER JOIN QMaxCourtDates ON [CourtDates].ID=[QMaxCourtDates].CourtDatesID;

SELECT ID, (SELECT Sum(GroupTasksIncomplete.TimeLength) AS Total FROM GroupTasksIncomplete WHERE GroupTasksIncomplete.ID <= T1.ID) AS Total
FROM GroupTasksIncomplete AS T1;

SELECT CourtDates.ID AS CourtDatesID, CourtDates.InvoiceNo AS OAIInvoiceNo, CourtDates.Subtotal AS OAISubtotal, CourtDates.Quantity AS OAIQuantity, CourtDates.UnitPrice AS OAIUnitPrice, CourtDates.PaymentSum AS PaymentSum, Customers.ID AS CustomersID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.FaxNumber, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.Notes, Customers.FactoringApproved, CourtDates.CasesID AS OAICasesID
FROM Customers INNER JOIN CourtDates ON Customers.ID = CourtDates.OrderingID
WHERE (((CourtDates.ID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT Tasks.[ID], Tasks.[Title], Tasks.[Priority], Tasks.[Description], Tasks.[Due Date], Tasks.[PriorityPoints], Tasks.[Category], Tasks.[TimeLength], Tasks.[CourtDatesID]
FROM Tasks
WHERE Tasks.Completed=No
GROUP BY Tasks.CourtDatesID, Tasks.Priority, Tasks.Title, Tasks.PriorityPoints, Tasks.DueDate;

SELECT *
FROM ShippingOptionsQ
WHERE [id] = 1;

SELECT DSum([CourtDates].[Subtotal],"PaymentQueryInvoiceInfo") AS FinalPrice, Payments.ID AS PaymentsID, Payments.InvoiceNo AS pInvoiceNo, Payments.Amount, Payments.RemitDate, CourtDates.ID AS CourtDatesID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.DueDate, CourtDates.InvoiceNo AS cInvoiceNo, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.Subtotal, CourtDates.UnitPrice
FROM Payments INNER JOIN CourtDates ON Payments.InvoiceNo = CourtDates.InvoiceNo
WHERE (((CourtDates.ID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT Payments.ID AS ["PaymentsID"], Payments.InvoiceNo AS ["PaymentsInvoiceNo"], Payments.Amount AS ["Amount"], Payments.RemitDate AS ["RemitDate"], CourtDates.ID AS ["CourtDatesID"], CourtDates.HearingDate AS ["HearingDate"], CourtDates.HearingStartTime AS ["HearingStartTime"], CourtDates.HearingEndTime AS ["HearingEndTime"], CourtDates.CasesID AS ["CasesID"], CourtDates.OrderingID AS ["OrderingID"], CourtDates.AudioLength AS ["AudioLength"], CourtDates.TurnaroundTimesCD AS ["TurnaroundTimesCD"], CourtDates.InvoiceNo AS ["InvoiceNo"], CourtDates.InvoiceDate AS ["InvoiceDate"], CourtDates.PaymentDueDate AS ["PaymentDueDate"], CourtDates.UnitPrice AS ["UnitPrice"], CourtDates.Quantity AS ["Quantity"], CourtDates.Subtotal AS ["Subtotal"]
FROM Payments INNER JOIN CourtDates ON Payments.InvoiceNo=CourtDates.InvoiceNo;

SELECT CourtDates.ID AS ["CourtDatesID"], CourtDates.HearingDate, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.DueDate, CourtDates.InvoiceNo AS ["CourtDatesInvoiceNo"], CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.Subtotal, Payments.InvoiceNo AS ["PaymentsInvoiceNo"], Payments.Amount, Payments.RemitDate, Payments.ID AS ["PaymentsID"]
FROM CourtDates INNER JOIN Payments ON CourtDates.[InvoiceNo] = Payments.[InvoiceNo];

SELECT Payments.InvoiceNo AS pInvoiceNo, CourtDates.ID AS CourtDatesID, CourtDates.UnitPrice, CourtDates.ActualQuantity, CourtDates.AudioLength, CourtDates.InvoiceDate, DSum([CourtDates].[Subtotal],"PaymentQueryInvoiceInfo3") AS FinalPrice, Payments.ID AS PaymentsID, Payments.Amount, Payments.RemitDate, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.TurnaroundTimesCD, CourtDates.DueDate, CourtDates.InvoiceNo AS cInvoiceNo, CourtDates.PaymentDueDate, CourtDates.Subtotal
FROM Payments INNER JOIN CourtDates ON Payments.InvoiceNo = CourtDates.InvoiceNo
WHERE (((Payments.InvoiceNo)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![InvoiceNo]));

SELECT InvoicesQuery4.EmailAddress AS [Recipient Email], InvoicesQuery4.FirstName AS [Recipient First Name], InvoicesQuery4.LastName AS [Recipient Last Name], InvoicesQuery4.InvoiceNo AS [Invoice Number], (DateAdd("d", 1, Date())) AS [Due Date], InvoicesQuery4.CourtDatesID AS Reference, CourtDatesRatesQuery.Code AS [Item Name], "|   " & InvoicesQuery4.Party1 & "  v. " & InvoicesQuery4.Party2 & "   |"   & Chr(13) & "|   " & InvoicesQuery4.CaseNumber1 & "   " & InvoicesQuery4.CaseNumber2 & "   |   Hearing Date:  " & InvoicesQuery4.HearingDate & "   |" & Chr(13) & "|   Approx. " & InvoicesQuery4.AudioLength & " minutes   |   " & InvoicesQuery4.TurnaroundTimesCD & " calendar-day turnaround   |" & Chr(13) & "   |" AS Description, InvoicesQuery4.Quantity AS [Item Amount], "" AS [Shipping Cost], "" AS [Discount Amount], "USD" AS [Currency Code], "Once both audio and a deposit has been received, the turnaround time will begin.  We will complete the transcript.  After transcript completion and final payment, the transcript will be filed if applicable as well as e-mailed to you in Word and PDF versions.  We will upload it to our online repository for your 24/7 access.  Two copies are included in our rate.  If we are filing this with the Court of Appeals, one is mailed to the court and the other to you.  Otherwise, you will receive both copies.  Our transcripts also include a weatherproof color-labeled CD of your audio and transcript.  If you don't want the hard copies mailed or just want the CD, that's fine, too; just let us know.  If this is filed with the Court of Appeals, you will receive a notification upon filing directly from the court.  If I have any spellings questions or things like that (hopefully not), I will let you know." AS [Note to Customer], "This is an invoice for deposit.  The deposit amount has been calculated as 100 percent of the estimated cost of the transcript.  The balance remaining will be due/refunded upon completion of the transcript after a final page count has been determined.  Please check out our full terms of service at http://www.aquoco.co/ServiceA.html.  Thank you for your business." AS [Terms and Conditions], InvoicesQuery4.CourtDatesID AS [Memo to Self]
FROM InvoicesQuery4 INNER JOIN CourtDatesRatesQuery ON CourtDatesRatesQuery.CourtDatesID=InvoicesQuery4.CourtDatesID
WHERE (InvoicesQuery4.CourtDatesID=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]);

SELECT InvoicesQuery4.EmailAddress AS [Recipient Email], InvoicesQuery4.FirstName AS [Recipient First Name], InvoicesQuery4.LastName AS [Recipient Last Name], InvoicesQuery4.InvoiceNo AS [Invoice Number], (DateAdd("d", 1, Date())) AS [Due Date], InvoicesQuery4.CourtDatesID AS Reference, CourtDatesRatesQuery.Code AS [Item Name], "|   " & InvoicesQuery4.Party1 & "  v. " & InvoicesQuery4.Party2 & "   |"   & Chr(13) & "|   " & InvoicesQuery4.CaseNumber1 & "   " & InvoicesQuery4.CaseNumber2 & "   |   Hearing Date:  " & InvoicesQuery4.HearingDate & "   |" & Chr(13) & "|   Approx. " & InvoicesQuery4.AudioLength & " minutes   |   " & InvoicesQuery4.TurnaroundTimesCD & " calendar-day turnaround   |" & Chr(13) & "   |" AS Description, InvoicesQuery4.Quantity AS [Item Amount], "" AS [Shipping Cost], "" AS [Discount Amount], "USD" AS [Currency Code], "This is an order confirmation and estimated price quote for the work you requested.  The details of your request and due date is listed on this quote for your convenience.  In terms of next steps, once audio has been received, the turnaround time will begin.  We will complete the transcript.  After transcript completion, the transcript will be filed if applicable as well as e-mailed to you in Word and PDF versions.  We will upload the transcript to our online repository for your 24/7 access.    You will receive an invoice at the time of completion.  Two copies are included in our rate.  If we are filing this with the Court of Appeals, one is mailed to the court and the other to you.  Otherwise, you will receive both copies.  Our transcripts also include a weatherproof color-labeled CD of your audio and transcript.  If you don't want the hard copies mailed or just want the CD, that's fine, too; just let us know.  Otherwise, we will just mail out as described previously.  If this is filed with the Court of Appeals, you will receive a notification upon filing directly from the court." AS [Note to Customer], "Please pay within 28 days.  5% interest if payment received after 28 calendar days of invoice date, additional 1% interest added every 7th calendar day after day 28 up to a maximum of 12%.  Please submit payment to A Quo Co., c/o American Funding Solutions, PO Box 572, Blue Springs, MO 64013.  Please check out our full terms of service at http://www.aquoco.co/ServiceA.html.  Thank you for your business." AS [Terms and Conditions], InvoicesQuery4.CourtDatesID AS [Memo to Self]
FROM InvoicesQuery4 INNER JOIN CourtDatesRatesQuery ON CourtDatesRatesQuery.CourtDatesID=InvoicesQuery4.CourtDatesID
WHERE (InvoicesQuery4.CourtDatesID=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]);

SELECT *
FROM Statuses
WHERE ((Statuses.[CourtDatesID])=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]);

SELECT FinalUnitPriceInvoiceQuery.CourtDatesID AS CourtDatesID, (QTotalPricebyInvoiceNumber.[TotalPrice]-[QTotalPaymentsbyInvoiceNumber].[TotalPayments]) AS BalanceOwed, FinalUnitPriceInvoiceQuery.*, QTotalPaymentsbyInvoiceNumber.TotalPayments, QTotalFactoringCostbyInvoiceNumber.TotalFactoringCost, QTotalPricebyLastFirstName.TotalPrice
FROM ((FinalUnitPriceInvoiceQuery LEFT JOIN QTotalPricebyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalPricebyInvoiceNumber.InvoiceNo) LEFT JOIN QTotalPaymentsbyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalPaymentsbyInvoiceNumber.InvoiceNo) LEFT JOIN QTotalFactoringCostbyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalFactoringCostbyInvoiceNumber.InvoiceNo
WHERE ((((QTotalPricebyInvoiceNumber.[TotalPrice]-[QTotalPaymentsbyInvoiceNumber].[TotalPayments]))>0));

SELECT FindReplaceShortcuts.ID, FindReplaceShortcuts.Find, FindReplaceShortcuts.BankruptcyReplace
FROM FindReplaceShortcuts;

SELECT CourtDates.ID AS CourtDatesID, CourtDates.CasesID AS CasesID, Cases.Jurisdiction, Doctors.LX, Doctors.L1, Doctors.L2, Doctors.L3, Doctors.L4, Doctors.L5, Doctors.L6, Doctors.L7, Doctors.L8, Doctors.L9, Doctors.L10, Doctors.L11, Doctors.L12, Doctors.L13, Doctors.L14, Doctors.L15, Doctors.L16, Doctors.L17, Doctors.L18, Doctors.L19, Doctors.L20, Doctors.L21, Doctors.L22, Doctors.L23, Doctors.L24, Doctors.L25, Doctors.L26, Doctors.L27, Doctors.L28, Doctors.L29, Doctors.L30, Doctors.L31, Doctors.L32, Doctors.L33, Doctors.L34, Doctors.L35, Doctors.L36, Doctors.L37, Doctors.L38, Doctors.L39, Doctors.L40, Doctors.L41, Doctors.L42, Doctors.L43, Doctors.L44, Doctors.L45, Doctors.L46, Doctors.L47, Doctors.L48, Doctors.L49, Doctors.L50, Doctors.L51, Doctors.L52, Doctors.L53, Doctors.L54, Doctors.L55, Doctors.L56, Doctors.L57, Doctors.L58, Doctors.L59, Doctors.L60, Doctors.L61, Doctors.L62, Doctors.L63, Doctors.L64, Doctors.L65, Doctors.L66, Doctors.L67, Doctors.L68, Doctors.L69, Doctors.L70
FROM (Cases INNER JOIN Doctors ON (Cases.[Jurisdiction] = Doctors.[Jurisdiction]) AND (Cases.[ID] = Doctors.[CasesID])) INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]
WHERE (((CourtDates.ID)=(Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField])));

SELECT [SubtotalUnitPriceQuery].InvoiceNo, [SubtotalUnitPriceQuery].InvoiceDate, CDbl(Nz(Sum([SubtotalUnitPriceQuery].[Quantity]),0)) AS PageCount, CDbl(Nz(Sum([SubtotalUnitPriceQuery].[Subtotal]),2)) AS Subtotal, CDbl(Nz(Sum([SubtotalUnitPriceQuery].[AudioLength]),0)) AS AudioLength
FROM SubtotalUnitPriceQuery
GROUP BY [SubtotalUnitPriceQuery].InvoiceNo, [SubtotalUnitPriceQuery].InvoiceDate;

SELECT CourtDates.ID AS CourtdatesID, CourtDates.InvoiceNo AS InvoiceNo, Customers.ID AS CustomersID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.FaxNumber, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.Notes, Customers.FactoringApproved, CourtDates.CasesID
FROM Customers INNER JOIN CourtDates ON CourtDates.OrderingID = Customers.ID
WHERE CourtDates.ID = [CourtDatesID];

SELECT FindReplaceShortcuts.JEWFDA1, FindReplaceShortcuts.ID
FROM FindReplaceShortcuts;

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.UnitPrice, CourtDates.Quantity, CourtDates.Subtotal, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.DueDate, CourtDates.InvoiceDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.HearingDate, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.FiledNotFiled, CourtDates.PaymentDueDate, CourtDates.ExpectedAdvanceDate, CourtDates.ExpectedRebateDate, CourtDates.EstimatedPageCount, Cases.Party1, Cases.Party2, Cases.Party1Name, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge
FROM CourtDates INNER JOIN CASES ON CourtDates.CasesID = Cases.ID
WHERE CourtDates.ID=Forms![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField];

SELECT *
FROM TRInvoiceCasesQ INNER JOIN TRAppAddrInvQ ON [TRInvoiceCasesQ].[OrderingID]=[TRAppAddrInvQ].[ID];

SELECT Customers.ID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.FactoringApproved
FROM Customers INNER JOIN InvoiceInfoQ ON Customers.ID = InvoiceInfoQ.ID
WHERE (((Customers.ID)=[InvoiceInfoQ].[OrderingID]));

SELECT Max(CourtDates.ID) AS CourtDatesID
FROM Courtdates;

SELECT QInfobyInvoiceNumber.ID AS QInfobyInvoiceNumber_ID, QInfobyInvoiceNumber.Party1, QInfobyInvoiceNumber.Party2, QInfobyInvoiceNumber.Party1Name, QInfobyInvoiceNumber.Party2Name, QInfobyInvoiceNumber.CaseNumber1, QInfobyInvoiceNumber.CaseNumber2, QInfobyInvoiceNumber.Jurisdiction, QInfobyInvoiceNumber.HearingTitle, QInfobyInvoiceNumber.Judge, QInfobyInvoiceNumber.InvoiceNo, QInfobyInvoiceNumber.UnitPrice, QInfobyInvoiceNumber.Quantity, QInfobyInvoiceNumber.Subtotal, QInfobyInvoiceNumber.AudioLength, QInfobyInvoiceNumber.TurnaroundTimesCD, QInfobyInvoiceNumber.DueDate, QInfobyInvoiceNumber.InvoiceDate, QInfobyInvoiceNumber.HearingStartTime, QInfobyInvoiceNumber.HearingEndTime, QInfobyInvoiceNumber.HearingDate, QInfobyInvoiceNumber.CasesID, QInfobyInvoiceNumber.OrderingID, QInfobyInvoiceNumber.FiledNotFiled, QInfobyInvoiceNumber.PaymentDueDate, QInfobyInvoiceNumber.ExpectedAdvanceDate, QInfobyInvoiceNumber.ExpectedRebateDate, QInfobyInvoiceNumber.EstimatedPageCount, Customers.ID AS Customers_ID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.JobTitle, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.FactoringApproved, (QInfobyInvoiceNumber.[Subtotal]*.8) AS ExpectedAdvanceAmount
FROM Customers INNER JOIN QInfobyInvoiceNumber ON Customers.[ID] = QInfobyInvoiceNumber.[OrderingID];

SELECT CourtDates.ID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.Location, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, CourtDates.DueDate, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.PaymentType, CourtDates.Notes, CourtDates.ShippingOptionsID, CourtDates.SPKRID, CourtDates.AGShortcuts, CourtDates.FiledNotFiled, CourtDates.Factored, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.FactoringInterestID, CourtDates.ExpectedRebateDate, CourtDates.EstimatedPageCount, CourtDates.FactoringCost, CourtDates.UnitPrice, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.Subtotal, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.PaymentSum, CourtDates.InventoryRateCode, CourtDates.AccountCode, CourtDates.TaxType, CourtDates.BrandingTheme, CourtDates.PPID, CourtDates.PPStatus, CourtDates.FinalPrice AS FinalPrice, CourtDates.ID AS CourtDatesID, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.Subtotal, CourtDates.UnitPrice, CourtDates.ActualQuantity, CourtDates.PaymentSum, CourtDates.PPStatus, CourtDates.PPID, Customers.Company, Customers.FirstName, Customers.LastName, Customers.Address, Customers.City, Customers.State, Customers.Zip, Customers.FactoringApproved
FROM CourtDates INNER JOIN Customers ON CourtDates.OrderingID = Customers.ID
WHERE (((CourtDates.PPStatus)<>'PAID') or (CourtDates.PPStatus<>'MARKED_AS_PAID'));

SELECT DISTINCT CourtDates.OrderingID, Customers.Company, Customers.FirstName, Customers.LastName, CourtDates.InvoiceNo AS cInvoiceNo, CourtDates.InvoiceDate AS cInvoiceDate, QEstimatedPricebyInvoiceNumber.InvoiceNo AS qInvoiceNo, QEstimatedPricebyInvoiceNumber.Subtotal AS qSubtotal, QEstimatedPricebyInvoiceNumber.PageCount AS qPageCount, QEstimatedPricebyInvoiceNumber.AudioLength AS qAudioLength
FROM Customers INNER JOIN (QEstimatedPricebyInvoiceNumber INNER JOIN CourtDates ON QEstimatedPricebyInvoiceNumber.[InvoiceNo] = CourtDates.[InvoiceNo]) ON Customers.[ID] = CourtDates.[OrderingID]
WHERE (((QEstimatedPricebyInvoiceNumber.InvoiceNo)=[CourtDates].[InvoiceNo]))
GROUP BY QEstimatedPricebyInvoiceNumber.InvoiceNo, CourtDates.InvoiceNo, CourtDates.OrderingID, Customers.Company, Customers.FirstName, Customers.LastName, CourtDates.InvoiceDate, QEstimatedPricebyInvoiceNumber.Subtotal, QEstimatedPricebyInvoiceNumber.PageCount, QEstimatedPricebyInvoiceNumber.AudioLength;

SELECT *
FROM TempCourtDates;

SELECT Title, PriorityPoints, [Due Date], TimeLength, Description, Completed
FROM Tasks
WHERE [Title] Like '*' & Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField & '*' AND [Completed]=False;

SELECT Expenses.Memo, Expenses.ExpensesDate, CDbl(Nz(Sum([Expenses].[Amount]),0)) AS TotalExpenses
FROM Expenses
GROUP BY Expenses.Memo, Expenses.ExpensesDate;

SELECT CourtDates.InvoiceNo, QTotalExpensesbyInvoiceNumber.InvoiceNo
FROM CourtDates INNER JOIN QTotalExpensesbyInvoiceNumber ON CourtDates.InvoiceNo <> QTotalExpensesbyInvoiceNumber.InvoiceNo;

SELECT [Expenses].InvoiceNo, CDbl(Nz(Sum([Expenses].[Amount]),0)) AS TotalExpenses
FROM Expenses
GROUP BY [Expenses].InvoiceNo;

SELECT InvoiceNo, TotalExpenses
FROM QTotalInvoicesNoExpenses
UNION SELECT InvoiceNo, TotalExpenses
FROM QTotalExpensesbyInvoiceNumber;

SELECT [Expenses].CourtDatesID, CDbl(Nz(Sum([Expenses].[Amount]),0)) AS TotalExpenses
FROM Expenses
GROUP BY [Expenses].CourtDatesID;

SELECT [FinalUnitPriceQuery].InvoiceNo, CDbl(Nz(Sum([FinalUnitPriceQuery].[FactoringCost]),0)) AS TotalFactoringCost
FROM FinalUnitPriceQuery
GROUP BY [FinalUnitPriceQuery].InvoiceNo;

SELECT DISTINCT InvoiceNo, InvoiceDate
FROM CourtDates;

SELECT [CourtDatesWTOMatchingExp].InvoiceNo, CDbl(Nz(Sum([CourtDatesWTOMatchingExp].[TotalExpenses]),0)) AS TotalExpenses
FROM CourtDatesWTOMatchingExp
GROUP BY [CourtDatesWTOMatchingExp].InvoiceNo;

SELECT [CourtDates].InvoiceNo, CDbl(Nz(Sum([CourtDates].[ActualQuantity]),0)) AS PageCount
FROM CourtDates
GROUP BY [CourtDates].InvoiceNo;

SELECT [Payments].InvoiceNo, CDbl(Nz(Sum([Payments].[Amount]),0)) AS TotalPayments
FROM Payments
GROUP BY [Payments].InvoiceNo;

SELECT [FinalUnitPriceQuery].InvoiceNo, CDbl(Nz(Sum([FinalUnitPriceQuery].[ActualQuantity]),0)) AS PageCount, CDbl(Nz(Sum([FinalUnitPriceQuery].[FinalPrice]),2)) AS TotalPrice
FROM FinalUnitPriceQuery
GROUP BY [FinalUnitPriceQuery].InvoiceNo;

SELECT FinalUnitPriceInvoiceQuery.LastName, FinalUnitPriceInvoiceQuery.FirstName, FinalUnitPriceInvoiceQuery.InvoiceNo, CDbl(Nz(Sum([FinalUnitPriceInvoiceQuery].[FinalTotal]),0)) AS TotalPrice
FROM FinalUnitPriceInvoiceQuery
GROUP BY FinalUnitPriceInvoiceQuery.LastName, FinalUnitPriceInvoiceQuery.FirstName;

SELECT 
FROM CourtDates AS CourtDates_1, Statuses AS Statuses_1, Cases INNER JOIN (CourtDates INNER JOIN Statuses ON (Statuses.CourtDatesID = CourtDates.ID) AND (Statuses.ID = CourtDates.StatusesID) AND (CourtDates.StatusesID = Statuses.ID) AND (CourtDates.ID = Statuses.CourtDatesID)) ON (Statuses.ID = Cases.ID) AND (Cases.ID = CourtDates.CasesID);

SELECT MAX(ID) AS CourtDatesID, CourtDates.InvoiceNo
FROM CourtDates;

SELECT *
FROM Tasks
WHERE Priority='(2) Stage 2' AND Title LIKE '*1945*';

SELECT ID, CourtDatesID, PriorityPoints, [Due Date], Title, Description, Completed, Category, TimeLength
FROM Tasks
ORDER BY PriorityPoints DESC;

SELECT ID, CourtDatesID, PriorityPoints, [Due Date], Title, Description, Completed, Category, TimeLength, (SELECT Sum(runningtotaltasks.timelength) AS Total FROM runningtotaltasks WHERE runningtotaltasks.ID <= runningtotaltaskssum.ID) AS Total
FROM runningtotaltasks AS running
ORDER BY PriorityPoints DESC;

SELECT *
FROM Cases INNER JOIN CourtDates ON Cases.ID = CourtDates.CasesID
WHERE (((CourtDates.ID)=Forms!NewMainMenu!ProcessJobSubformNMM.Form!JobNumberField) And ((CourtDates.CasesID) Like Cases.ID));

SELECT *
FROM Statuses
WHERE ((Statuses.CourtDatesID)=(Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Customers.FirstName))like [Enter search term to search attorney's first name.  Enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((FirstName) Like [Enter search term to search ordering attorney's first name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, Cases.ID, CourtDates.CasesID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice, Cases.CaseNumber1
FROM (CourtDates INNER JOIN Cases ON CourtDates.CasesID = Cases.ID) INNER JOIN Customers ON CourtDates.OrderingID=Customers.ID
WHERE (((Cases.CaseNumber1) Like [Enter search term to search casenumber1; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, QBalanceOwed.Company, QBalanceOwed.MrMs, QBalanceOwed.LastName, QBalanceOwed.FirstName, QBalanceOwed.BusinessPhone, QBalanceOwed.Address, QBalanceOwed.City, QBalanceOwed.State, QBalanceOwed.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.EmailAddress, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE (((QBalanceOwed.[CaseNumber1]) Like [Enter search term to search casenumber1; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CitationHyperlinks.LongCitation, CitationHyperlinks.CHCategory, CitationHyperlinks.WebAddress, CitationHyperlinks.FindCitation, CitationHyperlinks.ReplaceHyperlink
FROM CitationHyperlinks
WHERE (((CitationHyperlinks.LongCitation)like [Enter search term to search citations; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, Cases.ID, CourtDates.CasesID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM Customers INNER JOIN (Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]) ON Customers.[ID] = CourtDates.[OrderingID]
WHERE (((Customers.Company) Like [Enter search term to search companies; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((Company) Like [Enter search term to search ordering company's name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Cases.Party2))like [Enter search term to search defendants; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((Party2) Like [Enter search term to search defendant's name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Customers.EmailAddress))like [Enter search term to search email addresses; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((EmailAddress) Like [Enter search term to search ordering company's email; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Cases.HearingTitle))like [Enter search term to search hearing titles; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((HearingTitle) Like [Enter search term to search hearing title; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((CourtDates.ID)like [Enter search term to search job numbers; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((CourtDatesID) Like [Enter search term to search job number; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((Judge) Like [Enter search term to search judge's name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM Customers, CourtDates INNER JOIN Cases ON CourtDates.CasesID = Cases.ID
WHERE (((Cases.Judge) Like [Enter search term to search judges; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Cases.Jurisdiction)like [Enter search term to search jurisdiction; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((Jurisdiction) Like [Enter search term to search by name of jurisdiction; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((LastName) Like [Enter search term to search ordering attorneys' last name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON Customers.ID = CourtDates.OrderingID) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Customers.Company)like [Enter search term to search ordering client's company; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON Customers.ID = CourtDates.OrderingID) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Customers.FirstName)like [Enter search term to search ordering attorney's first name; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Customers.LastName)like [Enter search term to search ordering attorney's last name; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT QBalanceOwed.InvoiceNo, QBalanceOwed.CourtDatesID, QBalanceOwed.FinalUnitPriceQuery_AudioLength AS AudioLength, QBalanceOwed.InvoiceInfoQ_InvoiceDate AS InvoiceDate, QBalanceOwed.InvoiceInfoQ_ActualQuantity AS FinalPageCount, QBalanceOwed.InvoiceInfoQ_ExpectedRebateDate AS ERebateDate, QBalanceOwed.InvoiceInfoQ_ExpectedAdvanceDate AS EAdvanceDate, QBalanceOwed.BalanceOwed, QBalanceOwed.FinalUnitPriceQuery_UnitPrice AS PageRate, QBalanceOwed.Party1, QBalanceOwed.Party2, QBalanceOwed.CaseNumber2, QBalanceOwed.HearingTitle, QBalanceOwed.Judge, QBalanceOwed.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, QBalanceOwed.HearingDate, QBalanceOwed.HearingStartTime, QBalanceOwed.HearingEndTime, QBalanceOwed.ShipDate, QBalanceOwed.TrackingNumber, QBalanceOwed.OrderingID, QBalanceOwed.CasesID, QBalanceOwed.Subtotal, QBalanceOwed.CaseNumber1, QBalanceOwed.FactoringCost
FROM QBalanceOwed
WHERE ((Party1) Like [Enter search term to search plaintiff's name; enter a * before and after to search with wildcard or it will search exact match:]);

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Cases.Party1) OR (Cases.Party2) like [Enter search term to search either party; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE (((Cases.Party1)like [Enter search term to search plaintiffs; enter a * before and after to search with wildcard or it will search exact match:]));

SELECT FinalUnitPriceInvoiceQuery.CourtDatesID AS CourtDatesID, (QTotalPricebyInvoiceNumber.[TotalPrice]-[QTotalPaymentsbyInvoiceNumber].[TotalPayments]) AS BalanceOwed, FinalUnitPriceInvoiceQuery.*, QTotalPaymentsbyInvoiceNumber.TotalPayments, QTotalFactoringCostbyInvoiceNumber.TotalFactoringCost, QTotalPricebyLastFirstName.TotalPrice
FROM ((FinalUnitPriceInvoiceQuery LEFT JOIN QTotalPricebyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalPricebyInvoiceNumber.InvoiceNo) LEFT JOIN QTotalPaymentsbyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalPaymentsbyInvoiceNumber.InvoiceNo) LEFT JOIN QTotalFactoringCostbyInvoiceNumber ON FinalUnitPriceInvoiceQuery.InvoiceNo = QTotalFactoringCostbyInvoiceNumber.InvoiceNo
WHERE ((((QTotalPricebyInvoiceNumber.TotalPrice-QTotalPaymentsbyInvoiceNumber.TotalPayments))>0) And FinalUnitPriceQuery_UnitPrice=40);

SELECT [Statuses].[ID]
FROM Statuses
WHERE ([Statuses].[ID] = Combo57);

SELECT ContactName, EmailAddress, POAddressLine1, POCity, PORegion, POPostalCode, InvoiceNumber, Reference, InvoiceDate, DueDate, InventoryItemCode, Description, Quantity, UnitAmount, AccountCode, TaxType, BrandingTheme
FROM XeroInvoiceCSV;

SELECT *
FROM ShippingOptions
WHERE [CourtDatesID] = Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField];

INSERT INTO ShippingOptions ( CourtDatesID )
SELECT [ShippingOptions]![CourtDatesID] AS Expr1, *
FROM ShippingOptions
WHERE ((([ShippingOptions]![CourtDatesID])=510));

SELECT SpeakersStatic.[ID], SpeakersStatic.[Jurisdiction], SpeakersStatic.[SPKR1], SpeakersStatic.[SPKR2], SpeakersStatic.[SPKR3], SpeakersStatic.[SPKR4], SpeakersStatic.[SPKR5], SpeakersStatic.[SPKR6], SpeakersStatic.[SPKR7], SpeakersStatic.[SPKR8], SpeakersStatic.[SPKR9], SpeakersStatic.[SPKR10], SpeakersStatic.[SPKR11], SpeakersStatic.[SPKR12], SpeakersStatic.[SPKR13], SpeakersStatic.[Spkr14], SpeakersStatic.[SPKR15], SpeakersStatic.[SPKR16], SpeakersStatic.[SPKR17], SpeakersStatic.[SPKR18], SpeakersStatic.[SPKR19], SpeakersStatic.[SPKR20], SpeakersStatic.[SPKR21], SpeakersStatic.[SPKR22], SpeakersStatic.[SPKR23], SpeakersStatic.[SPKR24], SpeakersStatic.[SPKR25], SpeakersStatic.[SPKR26], SpeakersStatic.[SPKR27], SpeakersStatic.[SPKR28], SpeakersStatic.[SPKR29], SpeakersStatic.[SPKR30], SpeakersStatic.[SPKR31], SpeakersStatic.[SPKR32], SpeakersStatic.[SPKR33], SpeakersStatic.[SPKR34], SpeakersStatic.[SPKR35], SpeakersStatic.[SPKR36], SpeakersStatic.[SPKR37], SpeakersStatic.[SPKR38], SpeakersStatic.[SPKR39], SpeakersStatic.[SPKR40], SpeakersStatic.[SPKR41], SpeakersStatic.[SPKR42], SpeakersStatic.[SPKR43], SpeakersStatic.[SPKR44], SpeakersStatic.[SPKR45], SpeakersStatic.[SPKR46], SpeakersStatic.[SPKR47], SpeakersStatic.[SPKR48], SpeakersStatic.[SPKR49], SpeakersStatic.[SPKR50], SpeakersStatic.[SPKR51], SpeakersStatic.[SPKR52], SpeakersStatic.[SPKR53], SpeakersStatic.[SPKR54], SpeakersStatic.[SPKR55], SpeakersStatic.[SPKR56], SpeakersStatic.[SPKR57], SpeakersStatic.[SPKR58], SpeakersStatic.[SPKR59], SpeakersStatic.[SPKR60], SpeakersStatic.[SPKR61], SpeakersStatic.[SPKR62], SpeakersStatic.[SPKR63], SpeakersStatic.[SPKR64], SpeakersStatic.[SPKR65], SpeakersStatic.[SPKR66], SpeakersStatic.[SPKR67], SpeakersStatic.[SPKR68], SpeakersStatic.[SPKR69], SpeakersStatic.[SPKR70], SpeakersStatic.[SPKR71], SpeakersStatic.[SPKR72], SpeakersStatic.[SPKR73], SpeakersStatic.[SPKR74], SpeakersStatic.[SPKR75], SpeakersStatic.[SPKR76], SpeakersStatic.[SPKR77], SpeakersStatic.[SPKR78], SpeakersStatic.[SPKR79], SpeakersStatic.[SPKR80], SpeakersStatic.[SPKR81], SpeakersStatic.[SPKR82], SpeakersStatic.[SPKR83], SpeakersStatic.[SPKR84], SpeakersStatic.[SPKR85], SpeakersStatic.[SPKR86], SpeakersStatic.[SPKR87], SpeakersStatic.[SPKR88], SpeakersStatic.[SPKR89], SpeakersStatic.[SPKR90], SpeakersStatic.[SPKR91], SpeakersStatic.[SPKR92], SpeakersStatic.[SPKR93], SpeakersStatic.[SPKR94], SpeakersStatic.[SPKR95], SpeakersStatic.[SPKR96], SpeakersStatic.[SPKR97], SpeakersStatic.[SPKR98], SpeakersStatic.[SPKR99], SpeakersStatic.[SPKR100]
FROM SpeakersStatic
WHERE SpeakersStatic.[ID]=2;

SELECT Cases.Notes, CourtDates.ID, CourtDates.CasesID, Cases.ID
FROM Cases INNER JOIN CourtDates ON CourtDates.[CasesID] = Cases.[ID]
WHERE ((CourtDates.ID Like Forms![MMProcess Jobs]!JobNumberField));

SELECT CourtDates.ID, CourtDates.InvoiceNo, CourtDates.AudioLength, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.Jurisdiction, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.Address, Customers.City, Customers.State, Customers.ZIP, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.Subtotal, CourtDates.ShipDate, CourtDates.TrackingNumber, CourtDates.InvoiceDate, CourtDates.Quantity, CourtDates.ActualQuantity, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.FinalPrice, CourtDates.UnitPrice
FROM (CourtDates INNER JOIN Customers ON (Customers.ID = CourtDates.OrderingID) OR (Customers.ID = CourtDates.App6) OR (Customers.ID = CourtDates.App5) OR (Customers.ID = CourtDates.App4) OR (Customers.ID = CourtDates.App3) OR (Customers.ID = CourtDates.App2) OR (Customers.ID = CourtDates.App1)) INNER JOIN Cases ON Cases.ID = CourtDates.CasesID
WHERE ((Customers.FirstName) like '*Ellis, Li*' OR (Customers.LastName) like '*Ellis, Li*' OR (Customers.Company) like '*Ellis, Li*' OR (Customers.EmailAddress) like '*Ellis, Li*' OR (Customers.BusinessPhone) like '*Ellis, Li*' OR (Customers.Address) like '*Ellis, Li*' OR (Customers.City) like '*Ellis, Li*' OR (Customers.State) like '*Ellis, Li*' OR (Customers.ZIP) like '*Ellis, Li*' OR (Cases.Party1) like '*Ellis, Li*' OR (Cases.Party1Name) like '*Ellis, Li*' OR (Cases.Party2) like '*Ellis, Li*' OR (Cases.Party2Name) like '*Ellis, Li*' OR (Cases.CaseNumber1) like '*Ellis, Li*' OR (Cases.CaseNumber2) like '*Ellis, Li*' OR (Cases.HearingTitle) like '*Ellis, Li*' OR (Cases.Judge) like '*Ellis, Li*' OR (Cases.JudgeTitle) like '*Ellis, Li*' OR (Cases.Jurisdiction) like '*Ellis, Li*' OR (Customers.Company) like '*Ellis, Li*' OR (CourtDates.HearingDate) like '*Ellis, Li*' OR (CourtDates.HearingStartTime) like '*Ellis, Li*' OR (CourtDates.HearingEndTime) like '*Ellis, Li*' OR (CourtDates.CasesID) like '*Ellis, Li*' OR (CourtDates.OrderingID) like '*Ellis, Li*' OR (CourtDates.Subtotal) like '*Ellis, Li*' OR (CourtDates.ShipDate) like '*Ellis, Li*' OR (CourtDates.TrackingNumber) like '*Ellis, Li*' OR (CourtDates.InvoiceDate) like '*Ellis, Li*' OR (CourtDates.Quantity) like '*Ellis, Li*' OR (CourtDates.ActualQuantity) like '*Ellis, Li*' OR (CourtDates.ExpectedRebateDate) like '*Ellis, Li*' OR (CourtDates.ExpectedAdvanceDate) like '*Ellis, Li*' OR (CourtDates.FinalPrice) like '*Ellis, Li*' OR (CourtDates.UnitPrice) like '*Ellis, Li*');

SELECT CourtDates.CasesID, Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.[CaseNumber1], Cases.[CaseNumber2], Cases.Jurisdiction, Cases.[HearingTitle], Cases.Judge, Cases.[JudgeTitle]
FROM Cases INNER JOIN CourtDates ON Cases.ID=CourtDates.CasesID
WHERE (((Cases.ID=CourtDates.CasesID) AND (CourtDates.ID) like Forms("[MMProcess Jobs]").Controls("ProcessJobSubform").Form.Controls("SCJSSBFM").Form.Controls("JobNumberField").Value));

SELECT Statuses.ContactsEntered, Statuses.JobEntered, Statuses.Stage1PpwkGenerated, Statuses.[Transcribe], Statuses.Stage3PpwkGenerated, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.Stage4PpwkGenerated, Statuses.Stage5PpwkGenerated, Statuses.BurnCD, Statuses.Mail, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber, Statuses.[CourtDatesID]
FROM Statuses INNER JOIN CourtDates ON (Statuses.[CourtDatesID])=(CourtDates.ID)
WHERE ((Statuses.CourtDatesID)=(CourtDates.ID));

SELECT CourtDates.ID AS CourtDatesID, UnitPrice.ID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, CourtDates.InvoiceDate AS InvoiceDate, CourtDates.PaymentDueDate, CourtDates.ExpectedAdvanceDate, CourtDates.ExpectedRebateDate, CourtDates.EstimatedPageCount, CourtDates.FactoringCost, CourtDates.UnitPrice, UnitPrice.Rate, CourtDates.Quantity AS Quantity, CourtDates.DueDate, CourtDates.Subtotal AS subSubtotal, Rate*Quantity AS Subtotal
FROM CourtDates INNER JOIN UnitPrice ON CourtDates.[UnitPrice] = UnitPrice.[ID]
WHERE CourtDates.FinalPrice = 0;

SELECT CourtDates.ID, Cases.Party1, Cases.Party2, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, CourtDates.AudioLength, CourtDates.DueDate, CourtDates.PaymentType, TaskMgmt.Hierarchy
FROM TaskMgmt INNER JOIN (Cases INNER JOIN CourtDates ON Cases.ID = CourtDates.CasesID) ON TaskMgmt.ID = Cases.ID
WHERE TaskMgmt.[CourtDatesID]=CourtDates.ID;

SELECT Tasks.[Due Date], Tasks.TimeLength, Tasks.Completed
FROM Tasks
WHERE (((Tasks.Completed)=False) AND ((Tasks.[Due Date])<=[Forms]![NewMainMenu]![ProcessJobSubformNMM]![SearchJobsSubform].[Form]![txtDeadline]));

SELECT ShippingOptionsQ.MailClass, ShippingOptionsQ.PackageType, ShippingOptionsQ.Width, ShippingOptionsQ.Length, ShippingOptionsQ.Depth, ShippingOptionsQ.PriorityMailExpress1030, ShippingOptionsQ.HolidayDelivery, ShippingOptionsQ.SundayDelivery, ShippingOptionsQ.SaturdayDelivery, ShippingOptionsQ.SignatureRequired, ShippingOptionsQ.Stealth, ShippingOptionsQ.ReplyPostage, ShippingOptionsQ.InsuredMail, ShippingOptionsQ.COD, ShippingOptionsQ.RestrictedDelivery, ShippingOptionsQ.AdultSignatureRestricted, ShippingOptionsQ.AdultSignatureRequired, ShippingOptionsQ.ReturnReceipt, ShippingOptionsQ.CertifiedMail, ShippingOptionsQ.SignatureConfirmation, ShippingOptionsQ.USPSTracking, ShippingOptionsQ.ReferenceID, ShippingOptionsQ.ToName, ShippingOptionsQ.ToAddress1, ShippingOptionsQ.ToAddress2, ShippingOptionsQ.ToCity, ShippingOptionsQ.ToState, ShippingOptionsQ.ToPostalCode, ShippingOptionsQ.ToCountry, ShippingOptionsQ.Value, ShippingOptionsQ.Description, ShippingOptionsQ.WeightOz, ShippingOptionsQ.ActualWeight, ShippingOptionsQ.ActualWeightText, ShippingOptionsQ.ID, ShippingOptionsQ.Output, ShippingOptionsQ.CourtDatesID
FROM ShippingOptionsQ
WHERE (((ShippingOptionsQ.CourtDatesID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT ShippingOptionsQ.MailClass, ShippingOptionsQ.PackageType, ShippingOptionsQ.Width, ShippingOptionsQ.Length, ShippingOptionsQ.Depth, ShippingOptionsQ.PriorityMailExpress1030, ShippingOptionsQ.HolidayDelivery, ShippingOptionsQ.SundayDelivery, ShippingOptionsQ.SaturdayDelivery, ShippingOptionsQ.SignatureRequired, ShippingOptionsQ.Stealth, ShippingOptionsQ.ReplyPostage, ShippingOptionsQ.InsuredMail, ShippingOptionsQ.COD, ShippingOptionsQ.RestrictedDelivery, ShippingOptionsQ.AdultSignatureRestricted, ShippingOptionsQ.AdultSignatureRequired, ShippingOptionsQ.ReturnReceipt, ShippingOptionsQ.CertifiedMail, ShippingOptionsQ.SignatureConfirmation, ShippingOptionsQ.USPSTracking, ShippingOptionsQ.ReferenceID, "Court of Appeals Div I Clerk's Office" AS ToName, "600 University St" AS ToAddress1, "One Union Square" AS ToAddress2, "Seattle" AS ToCity, "WA" AS ToState, "98101" AS ToPostalCode, ShippingOptionsQ.ToCountry, ShippingOptionsQ.Value, ShippingOptionsQ.Description, ShippingOptionsQ.WeightOz, ShippingOptionsQ.ActualWeight, ShippingOptionsQ.ActualWeightText, ShippingOptionsQ.ID, ShippingOptionsQ.Output, ShippingOptionsQ.CourtDatesID
FROM ShippingOptionsQ
WHERE (((ShippingOptionsQ.CourtDatesID)=[Forms]![NewMainMenu]![ProcessJobSubformNMM].[Form]![JobNumberField]));

SELECT Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.ID, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID
FROM Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID];

SELECT PaymentQueryInvoiceInfo.["PaymentsID"] AS Expr1, PaymentQueryInvoiceInfo.["PaymentsInvoiceNo"] AS Expr2, PaymentQueryInvoiceInfo.["Amount"] AS Expr3, PaymentQueryInvoiceInfo.["RemitDate"] AS Expr4, PaymentQueryInvoiceInfo.["CourtDatesID"] AS Expr5, PaymentQueryInvoiceInfo.["HearingDate"] AS Expr6, PaymentQueryInvoiceInfo.["HearingStartTime"] AS Expr7, PaymentQueryInvoiceInfo.["HearingEndTime"] AS Expr8, PaymentQueryInvoiceInfo.["CasesID"] AS Expr9, PaymentQueryInvoiceInfo.["OrderingID"] AS Expr10, PaymentQueryInvoiceInfo.["AudioLength"] AS Expr11, PaymentQueryInvoiceInfo.["TurnaroundTimesCD"] AS Expr12, PaymentQueryInvoiceInfo.["InvoiceNo"] AS Expr13, PaymentQueryInvoiceInfo.["InvoiceDate"] AS Expr14, PaymentQueryInvoiceInfo.["PaymentDueDate"] AS Expr15, PaymentQueryInvoiceInfo.["UnitPrice"] AS Expr16, PaymentQueryInvoiceInfo.["Quantity"] AS Expr17, PaymentQueryInvoiceInfo.["Subtotal"] AS Expr18
FROM PaymentQueryInvoiceInfo;

UPDATE TempShippingOptionsQ INNER JOIN ShippingOptions ON (TempShippingOptionsQ.PackageType = ShippingOptions.PackageType) AND (TempShippingOptionsQ.MailClass = ShippingOptions.MailClass) AND (TempShippingOptionsQ.CourtDatesID = ShippingOptions.CourtDatesID) AND (TempShippingOptionsQ.EmailAddress = ShippingOptions.ToEMail) AND (TempShippingOptionsQ.ZIP = ShippingOptions.ToPostalCode) AND (TempShippingOptionsQ.ToState = ShippingOptions.ToState) AND (TempShippingOptionsQ.ToCity = ShippingOptions.ToCity) AND (TempShippingOptionsQ.ToAddress1 = ShippingOptions.ToAddress2) AND (TempShippingOptionsQ.Company = ShippingOptions.ToAddress1) SET ShippingOptions.ToName = [TempShippingOptionsQ].[FirstName] & " " & [TempShippingOptionsQ].[LastName], ShippingOptions.ToAddress1 = [TempShippingOptionsQ]![Company], ShippingOptions.ToAddress2 = [TempShippingOptionsQ]![ToAddress1], ShippingOptions.ToCity = [TempShippingOptionsQ]![City], ShippingOptions.ToState = [TempShippingOptionsQ]![State], ShippingOptions.ToPostalCode = [TempShippingOptionsQ]![ZIP], ShippingOptions.ToEMail = [TempShippingOptionsQ]![EmailAddress];

UPDATE ShippingOptions INNER JOIN TempShippingOptionsQ ON TempShippingOptionsQ.CourtDatesID=ShippingOptions.CourtDatesID SET TempShippingOptionsQ.ToName = [ShippingOptions].[FirstName] & " " & [ShippingOptions].[LastName], TempShippingOptionsQ.ToAddress1 = [ShippingOptions]![Company], TempShippingOptionsQ.ToAddress2 = [ShippingOptions]![ToAddress1], TempShippingOptionsQ.ToCity = [ShippingOptions]![City], TempShippingOptionsQ.ToState = [ShippingOptions]![State], TempShippingOptionsQ.ToPostalCode = [ShippingOptions]![ZIP], TempShippingOptionsQ.ToEMail = [ShippingOptions]![EmailAddress];

SELECT #3/16/2020# AS Deadline, 1080 AS AudioLength, 810 AS PageCount, 2146.5 AS Subtotal1, 2632.5 AS Subtotal2, 3037.5 AS Subtotal3, 3442.5 AS Subtotal4, 2025 AS Subtotal5;

SELECT Customers.ID, Customers.MrMs, Customers.[Company], Customers.[LastName], Customers.[FirstName], Customers.[BusinessPhone], Customers.Address, Customers.City, Customers.[State], Customers.[ZIP], Customers.FactoringApproved, Customers.EmailAddress, Customers.Notes, GetInvoiceNoFromCDID.Subtotal*.8 AS EstimatedAdvanceAmount
FROM Customers INNER JOIN GetInvoiceNoFromCDID ON Customers.ID=[GetInvoiceNoFromCDID].OrderingID
WHERE Customers.ID=[GetInvoiceNoFromCDID].OrderingID;

SELECT Customers.ID, Customers.MrMs, Customers.[Company], Customers.[LastName], Customers.[FirstName], Customers.[BusinessPhone], Customers.Address, Customers.City, Customers.[State], Customers.[ZIP], Customers.FactoringApproved, Customers.EmailAddress, Customers.Notes
FROM Customers INNER JOIN [TR-Court-Q] ON (Customers.ID=[TR-Court-Q].OrderingID) OR (Customers.ID=[TR-Court-Q].App1) Or (Customers.ID=[TR-Court-Q].App2) Or (Customers.ID=[TR-Court-Q].App3) Or (Customers.ID=[TR-Court-Q].App4) Or (Customers.ID=[TR-Court-Q].App5) Or (Customers.ID=[TR-Court-Q].App6)
WHERE Customers.ID=[TR-Court-Q].OrderingID OR Customers.ID=[TR-Court-Q].App1 Or Customers.ID=[TR-Court-Q].App2 Or Customers.ID=[TR-Court-Q].App3 Or Customers.ID=[TR-Court-Q].App4 Or Customers.ID=[TR-Court-Q].App5 Or Customers.ID=[TR-Court-Q].App6;

SELECT Customers.ID, Customers.MrMs, Customers.[Company], Customers.[LastName], Customers.[FirstName], Customers.[BusinessPhone], Customers.Address, Customers.City, Customers.[State], Customers.[ZIP], Customers.FactoringApproved, Customers.EmailAddress, Customers.Notes
FROM Customers INNER JOIN [TR-Court-Q] ON (Customers.ID=[TR-Court-Q].OrderingID) OR (Customers.ID=[TR-Court-Q].App1) Or (Customers.ID=[TR-Court-Q].App2) Or (Customers.ID=[TR-Court-Q].App3) Or (Customers.ID=[TR-Court-Q].App4) Or (Customers.ID=[TR-Court-Q].App5) Or (Customers.ID=[TR-Court-Q].App6)
WHERE Customers.ID=[TR-Court-Q].OrderingID OR Customers.ID=[TR-Court-Q].App1 Or Customers.ID=[TR-Court-Q].App2 Or Customers.ID=[TR-Court-Q].App3 Or Customers.ID=[TR-Court-Q].App4 Or Customers.ID=[TR-Court-Q].App5 Or Customers.ID=[TR-Court-Q].App6;

SELECT Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.ID, CourtDates.Quantity, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTImesCD, CourtDates.PaymentDueDate, CourtDates.UnitPrice, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.Location, CourtDates.InvoiceNo, CourtDates.FactoringCost, CourtDates.InvoiceDate, CourtDates.Subtotal, CourtDates.FinalPrice, CourtDates.PaymentSum, CourtDates.EstimatedPageCount, CourtDates.DueDate, CourtDates.ActualQuantity, CourtDates.DueDate, CourtDates.InvoiceDate, CourtDates.FiledNotFiled, CourtDates.EstimatedPageCount, CourtDates.Location
FROM Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]
WHERE (((CourtDates.ID)=(Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField])));

SELECT Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.CourtDatesID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID
FROM Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID]
WHERE CourtDates.ID=([OrderingInfoForm]![HDTOrderingInfo]![Column(0)]);

SELECT Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.ID, CourtDates.Quantity, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTImesCD, CourtDates.PaymentDueDate, CourtDates.UnitPrice, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate, CourtDates.Location, CourtDates.InvoiceNo, CourtDates.FactoringCost, CourtDates.InvoiceDate, CourtDates.Subtotal, CourtDates.FinalPrice, CourtDates.PaymentSum, CourtDates.EstimatedPageCount, CourtDates.DueDate
FROM Cases INNER JOIN CourtDates ON Cases.[ID] = CourtDates.[CasesID];

SELECT Cases.Party1, Cases.Party1Name, Cases.Party2, Cases.Party2Name, Cases.CaseNumber1, Cases.CaseNumber2, Cases.Jurisdiction, Cases.HearingTitle, Cases.Judge, Cases.JudgeTitle, Cases.CourtDatesID
FROM Cases INNER JOIN CourtDates ON (Cases.ID = CourtDates.CasesID) AND (Cases.CourtDatesID = CourtDates.ID);

SELECT *
FROM [TR-AppAddrQ] INNER JOIN [TR-Court-Q] ON ([TR-AppAddrQ].ID=[TR-Court-Q].App6) Or ([TR-AppAddrQ].ID=[TR-Court-Q].App5) Or ([TR-AppAddrQ].ID=[TR-Court-Q].App4) Or ([TR-AppAddrQ].ID=[TR-Court-Q].App3) Or ([TR-AppAddrQ].ID=[TR-Court-Q].App2) Or ([TR-AppAddrQ].ID=[TR-Court-Q].App1) OR ([TR-AppAddrQ].ID=[TR-Court-Q].OrderingID);

SELECT CourtDates.ID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, Cases.Party1, Cases.Party2, Cases.Jurisdiction, Cases.HearingTitle, CourtDates.Location, CourtDates.CasesID
FROM Cases INNER JOIN CourtDates ON CourtDates.[CasesID]=Cases.[ID]
WHERE CourtDates.ID=(Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]);

SELECT CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.ID, CourtDates.Quantity, CourtDates.CasesID, CourtDates.App1, CourtDates.App2, CourtDates.App3, CourtDates.App4, CourtDates.App5, CourtDates.App6, CourtDates.OrderingID AS OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTImesCD, CourtDates.PaymentDueDate, CourtDates.UnitPrice, CourtDates.ExpectedRebateDate, CourtDates.ExpectedAdvanceDate AS ExpectedAdvanceDate, CourtDates.Location, CourtDates.InvoiceNo, CourtDates.FactoringCost, CourtDates.InvoiceDate, CourtDates.Subtotal, (CourtDates.Subtotal*.8) AS ExpectedAdvanceAmount, CourtDates.FinalPrice, CourtDates.PaymentSum, CourtDates.EstimatedPageCount, CourtDates.DueDate, CourtDates.ActualQuantity, CourtDates.DueDate, CourtDates.InvoiceDate, CourtDates.FiledNotFiled, CourtDates.EstimatedPageCount, CourtDates.PPID, CourtDates.PPStatus, GetInvoiceNoFromCDID.InvoiceNo, CourtDates.InventoryRateCode AS InventoryRateCode
FROM CourtDates INNER JOIN GetInvoiceNoFromCDID ON CourtDates.InvoiceNo=GetInvoiceNoFromCDID.InvoiceNo;

SELECT *
FROM Cases INNER JOIN GetInvoiceNoFromCDID ON Cases.ID=GetInvoiceNoFromCDID.OAICasesID;

SELECT *
FROM TRInv INNER JOIN TRAppAddrInvQ ON [TRInv].[CourtDates.OrderingID]=TRAppAddrInvQ.ID;

SELECT *
FROM TRInvoiceCasesQ INNER JOIN TRInv ON TRInvoiceCasesQ.[Cases.ID]=TRInv.CasesID;

SELECT TRInv.HearingDate AS TRInv_HearingDate, TRInv.HearingStartTime AS TRInv_HearingStartTime, TRInv.HearingEndTime AS TRInv_HearingEndTime, TRInv.ID, TRInv.Quantity AS TRInv_Quantity, TRInv.CasesID AS TRInv_CasesID, TRInv.App1 AS TRInv_App1, TRInv.App2 AS TRInv_App2, TRInv.App3 AS TRInv_App3, TRInv.App4 AS TRInv_App4, TRInv.App5 AS TRInv_App5, TRInv.App6 AS TRInv_App6, TRInv.OrderingID AS TRInv_OrderingID, TRInv.AudioLength AS TRInv_AudioLength, TRInv.TurnaroundTImesCD AS TRInv_TurnaroundTImesCD, TRInv.PaymentDueDate AS TRInv_PaymentDueDate, TRInv.UnitPrice AS TRInv_UnitPrice, TRInv.ExpectedRebateDate AS TRInv_ExpectedRebateDate, TRInv.ExpectedAdvanceDate AS TRInv_ExpectedAdvanceDate, TRInv.Location AS TRInv_Location, TRInv.CourtDates.InvoiceNo, TRInv.FactoringCost AS TRInv_FactoringCost, TRInv.Expr1022, TRInv.Subtotal AS TRInv_Subtotal, TRInv.ExpectedAdvanceAmount, TRInv.FinalPrice AS TRInv_FinalPrice, TRInv.PaymentSum, TRInv.Expr1027, TRInv.Expr1028, TRInv.ActualQuantity AS TRInv_ActualQuantity, TRInv.DueDate AS TRInv_DueDate, TRInv.InvoiceDate AS TRInv_InvoiceDate, TRInv.FiledNotFiled AS TRInv_FiledNotFiled, TRInv.EstimatedPageCount AS TRInv_EstimatedPageCount, TRInv.PPID AS TRInv_PPID, TRInv.PPStatus AS TRInv_PPStatus, TRInv.GetInvoiceNoFromCDID.InvoiceNo, TRInv.InventoryRateCode AS TRInv_InventoryRateCode, TRInvoiceCasesQ.Cases.ID, TRInvoiceCasesQ.Party1, TRInvoiceCasesQ.Party1Name, TRInvoiceCasesQ.Party2, TRInvoiceCasesQ.Party2Name, TRInvoiceCasesQ.CaseNumber1, TRInvoiceCasesQ.CaseNumber2, TRInvoiceCasesQ.Jurisdiction, TRInvoiceCasesQ.HearingTitle, TRInvoiceCasesQ.Judge, TRInvoiceCasesQ.JudgeTitle, TRInvoiceCasesQ.Cases.Notes, TRInvoiceCasesQ.GetInvoiceNoFromCDID.ID, TRInvoiceCasesQ.HearingDate AS TRInvoiceCasesQ_HearingDate, TRInvoiceCasesQ.HearingStartTime AS TRInvoiceCasesQ_HearingStartTime, TRInvoiceCasesQ.HearingEndTime AS TRInvoiceCasesQ_HearingEndTime, TRInvoiceCasesQ.CasesID AS TRInvoiceCasesQ_CasesID, TRInvoiceCasesQ.App1 AS TRInvoiceCasesQ_App1, TRInvoiceCasesQ.App2 AS TRInvoiceCasesQ_App2, TRInvoiceCasesQ.App3 AS TRInvoiceCasesQ_App3, TRInvoiceCasesQ.App4 AS TRInvoiceCasesQ_App4, TRInvoiceCasesQ.App5 AS TRInvoiceCasesQ_App5, TRInvoiceCasesQ.App6 AS TRInvoiceCasesQ_App6, TRInvoiceCasesQ.OrderingID AS TRInvoiceCasesQ_OrderingID, TRInvoiceCasesQ.StatusesID, TRInvoiceCasesQ.AudioLength AS TRInvoiceCasesQ_AudioLength, TRInvoiceCasesQ.Location AS TRInvoiceCasesQ_Location, TRInvoiceCasesQ.TurnaroundTimesCD AS TRInvoiceCasesQ_TurnaroundTimesCD, TRInvoiceCasesQ.InvoiceNo, TRInvoiceCasesQ.DueDate AS TRInvoiceCasesQ_DueDate, TRInvoiceCasesQ.ShipDate, TRInvoiceCasesQ.TrackingNumber, TRInvoiceCasesQ.PaymentType, TRInvoiceCasesQ.GetInvoiceNoFromCDID.CourtDates.Notes, TRInvoiceCasesQ.ShippingOptionsID, TRInvoiceCasesQ.SPKRID, TRInvoiceCasesQ.AGShortcuts, TRInvoiceCasesQ.FiledNotFiled AS TRInvoiceCasesQ_FiledNotFiled, TRInvoiceCasesQ.Factored, TRInvoiceCasesQ.InvoiceDate AS TRInvoiceCasesQ_InvoiceDate, TRInvoiceCasesQ.PaymentDueDate AS TRInvoiceCasesQ_PaymentDueDate, TRInvoiceCasesQ.FactoringInterestID, TRInvoiceCasesQ.ExpectedRebateDate AS TRInvoiceCasesQ_ExpectedRebateDate, TRInvoiceCasesQ.EstimatedPageCount AS TRInvoiceCasesQ_EstimatedPageCount, TRInvoiceCasesQ.FactoringCost AS TRInvoiceCasesQ_FactoringCost, TRInvoiceCasesQ.UnitPrice AS TRInvoiceCasesQ_UnitPrice, TRInvoiceCasesQ.Quantity AS TRInvoiceCasesQ_Quantity, TRInvoiceCasesQ.ActualQuantity AS TRInvoiceCasesQ_ActualQuantity, TRInvoiceCasesQ.Subtotal AS TRInvoiceCasesQ_Subtotal, TRInvoiceCasesQ.ExpectedAdvanceDate AS TRInvoiceCasesQ_ExpectedAdvanceDate, TRInvoiceCasesQ.FinalPrice AS TRInvoiceCasesQ_FinalPrice, TRInvoiceCasesQ.GetInvoiceNoFromCDID.CourtDates.PaymentSum, TRInvoiceCasesQ.InventoryRateCode AS TRInvoiceCasesQ_InventoryRateCode, TRInvoiceCasesQ.AccountCode, TRInvoiceCasesQ.TaxType, TRInvoiceCasesQ.BrandingTheme, TRInvoiceCasesQ.PPID AS TRInvoiceCasesQ_PPID, TRInvoiceCasesQ.PPStatus AS TRInvoiceCasesQ_PPStatus, TRInvoiceCasesQ.CourtDatesID AS TRInvoiceCasesQ_CourtDatesID, TRInvoiceCasesQ.OAIInvoiceNo, TRInvoiceCasesQ.OAISubtotal, TRInvoiceCasesQ.OAIQuantity, TRInvoiceCasesQ.OAIUnitPrice, TRInvoiceCasesQ.OrderingAttorneyInfo.PaymentSum, TRInvoiceCasesQ.CustomersID, TRInvoiceCasesQ.Company, TRInvoiceCasesQ.MrMs, TRInvoiceCasesQ.LastName, TRInvoiceCasesQ.FirstName, TRInvoiceCasesQ.EmailAddress, TRInvoiceCasesQ.BusinessPhone, TRInvoiceCasesQ.FaxNumber, TRInvoiceCasesQ.Address, TRInvoiceCasesQ.City, TRInvoiceCasesQ.State, TRInvoiceCasesQ.ZIP, TRInvoiceCasesQ.OrderingAttorneyInfo.Notes, TRInvoiceCasesQ.FactoringApproved, TRInvoiceCasesQ.OAICasesID, CourtDatesRatesQuery.CourtDatesID AS CourtDatesRatesQuery_CourtDatesID, CourtDatesRatesQuery.InventoryRateCode AS CourtDatesRatesQuery_InventoryRateCode, CourtDatesRatesQuery.[List Price], CourtDatesRatesQuery.RatesID, CourtDatesRatesQuery.Code
FROM CourtDatesRatesQuery INNER JOIN (TRInvoiceCasesQ INNER JOIN TRInv ON TRInvoiceCasesQ.[CustomersID] = TRInv.[OrderingID]) ON CourtDatesRatesQuery.[CourtDatesID] = TRInv.[CourtDates.ID];

SELECT TempCourtDates.[CourtDatesID], UnitPrice.[ID], TempCourtDates.[AudioLength], TempCourtDates.TurnaroundTimesCD, TempCourtDates.InvoiceNo, TempCourtDates.InvoiceDate, TempCourtDates.EstimatedPageCount, TempCourtDates.Quantity, TempCourtDates.DueDate, TempCourtDates.UnitPrice, UnitPrice.Rate, Rate*Quantity AS ["Subtotal"], (DateAdd('d',30,DueDate)) AS ["ExpectedRebateDate"], (DateAdd('d',2,[DueDate])) AS ["ExpectedAdvanceDate"]
FROM TempCourtDates INNER JOIN UnitPrice ON TempCourtDates.[UnitPrice] = UnitPrice.[ID];

SELECT *
FROM Statuses
WHERE (((Statuses.CourtDatesID)=Forms![SBFMUncompletedStatuses]![Combo57]));

SELECT Statuses.ContactsEntered, Statuses.JobEntered, Statuses.CoverPage, Statuses.AutoCorrect, Statuses.Schedule, Statuses.Invoice, Statuses.Transcribe, Statuses.AddRDtoCover, Statuses.FindReplaceRD, Statuses.HyperlinkTranscripts, Statuses.SpellingsEmail, Statuses.AudioProof, Statuses.InvoiceCompleted, Statuses.NoticeofService, Statuses.PackageEnclosedLetter, Statuses.CDLabel, Statuses.GenerateZIPs, Statuses.TranscriptsReady, Statuses.InvoicetoFactorEmail, Statuses.FileTranscript, Statuses.BurnCD, Statuses.ShippingXMLs, Statuses.GenerateShippingEM, Statuses.AddTrackingNumber, Statuses.CourtDatesID
FROM Statuses
WHERE Statuses.ContactsEntered=0 OR Statuses.JobEntered=0 OR Statuses.CoverPage=0 OR Statuses.AutoCorrect=0 OR Statuses.Schedule=0 OR Statuses.Invoice=0 OR Statuses.Transcribe=0 OR Statuses.AddRDtoCover=0 OR Statuses.FindReplaceRD=0 OR Statuses.HyperlinkTranscripts=0 OR Statuses.SpellingsEmail=0 OR Statuses.AudioProof=0 OR Statuses.InvoiceCompleted=0 OR Statuses.NoticeofService=0 OR Statuses.PackageEnclosedLetter=0 OR Statuses.CDLabel=0 OR Statuses.GenerateZIPs=0 OR Statuses.TranscriptsReady=0 OR Statuses.InvoicetoFactorEmail=0 OR Statuses.FileTranscript=0 OR Statuses.BurnCD=0 OR Statuses.ShippingXMLs=0 OR Statuses.GenerateShippingEM=0 OR Statuses.AddTrackingNumber=0;

SELECT CourtDates.ID AS CourtDatesID, UnitPrice.ID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.StatusesID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.InvoiceNo, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.EstimatedPageCount, CourtDates.Quantity, CourtDates.DueDate, CourtDates.UnitPrice, UnitPrice.Rate, UnitPrice.Rate*CourtDates.Quantity AS Subtotal, (DateAdd('d',30,DueDate)) AS ExpectedRebateDate, (DateAdd('d',2,[DueDate])) AS ExpectedAdvanceDate, Subtotal*.8 AS EstimatedAdvanceAmount
FROM CourtDates INNER JOIN UnitPrice ON CourtDates.[UnitPrice] = UnitPrice.[ID];

SELECT [CourtDates].[FinalPrice] AS FinalPrice, Payments.ID AS PaymentsID, Payments.InvoiceNo AS pInvoiceNo, Payments.Amount, Payments.RemitDate, CourtDates.ID AS CourtDatesID, CourtDates.HearingDate, CourtDates.HearingStartTime, CourtDates.HearingEndTime, CourtDates.CasesID, CourtDates.OrderingID, CourtDates.AudioLength, CourtDates.TurnaroundTimesCD, CourtDates.DueDate, CourtDates.InvoiceNo AS cInvoiceNo, CourtDates.InvoiceDate, CourtDates.PaymentDueDate, CourtDates.Subtotal, CourtDates.UnitPrice, CourtDates.ActualQuantity, CourtDates.PaymentSum, Customers.Company, Customers.FirstName, Customers.LastName, Customers.Address, Customers.City, Customers.State, Customers.Zip, Customers.FactoringApproved
FROM (Payments INNER JOIN CourtDates ON Payments.InvoiceNo = CourtDates.InvoiceNo) INNER JOIN Customers ON CourtDates.OrderingID = Customers.ID
WHERE ((CourtDates.[FinalPrice]-CourtDates.[PaymentSum])>1);

SELECT UnpaidInvoicesQ.[FinalPrice], UnpaidInvoicesQ.[PaymentsID], UnpaidInvoicesQ.[pInvoiceNo], UnpaidInvoicesQ.[Amount], UnpaidInvoicesQ.[RemitDate], UnpaidInvoicesQ.[CourtDatesID], UnpaidInvoicesQ.[HearingDate], UnpaidInvoicesQ.[HearingStartTime], UnpaidInvoicesQ.[HearingEndTime], UnpaidInvoicesQ.[CasesID], UnpaidInvoicesQ.[OrderingID], UnpaidInvoicesQ.[AudioLength], UnpaidInvoicesQ.[TurnaroundTimesCD], UnpaidInvoicesQ.[DueDate], UnpaidInvoicesQ.[cInvoiceNo], UnpaidInvoicesQ.[InvoiceDate], UnpaidInvoicesQ.[PaymentDueDate], UnpaidInvoicesQ.[Subtotal], UnpaidInvoicesQ.[UnitPrice], UnpaidInvoicesQ.[ActualQuantity], UnpaidInvoicesQ.[PaymentSum], UnpaidInvoicesQ.[Company], UnpaidInvoicesQ.[FirstName], UnpaidInvoicesQ.[LastName], UnpaidInvoicesQ.[Address], UnpaidInvoicesQ.[City], UnpaidInvoicesQ.[State], UnpaidInvoicesQ.[Zip], UnpaidInvoicesQ.[FactoringApproved]
FROM UnpaidInvoicesQ
WHERE (((UnpaidInvoicesQ.[FactoringApproved])=True));

UPDATE CourtDates INNER JOIN InvoiceFPaymentDueDateQuery ON CourtDates.ID = InvoiceFPaymentDueDateQuery.CourtDatesID SET CourtDates.PaymentDueDate = ["PaymentDueDate"];

UPDATE CourtDates SET CourtDates.PaymentDueDate = (SELECT PaymentDueDate FROM InvoicePPaymentDueDateQuery WHERE ID = InvoicePPaymentDueDateQuery.CourtDatesID;);

SELECT CourtDates.ID, Customers.ID, Customers.Company, Customers.MrMs, Customers.LastName, Customers.FirstName, Customers.EmailAddress, Customers.BusinessPhone, Customers.FaxNumber, Customers.Address, Customers.City, Customers.State, Customers.ZIP, Customers.Notes, CourtDates.CasesID
FROM Customers INNER JOIN CourtDates ON (CourtDates.App1 = Customers.ID) OR (CourtDates.App2 = Customers.ID) OR (CourtDates.App3 = Customers.ID) OR (CourtDates.App4 = Customers.ID) OR (CourtDates.App5 = Customers.ID) OR (CourtDates.App6 = Customers.ID)
WHERE (CourtDates.ID=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField]);

SELECT BrandingThemes.ID, BrandingThemes.BrandingTheme AS BrandingThemes_BrandingTheme, XeroInvoiceCSV.BrandingTheme AS XeroInvoiceCSV_BrandingTheme
FROM XeroInvoiceCSV INNER JOIN BrandingThemes ON XeroInvoiceCSV.[BrandingTheme] = BrandingThemes.[BrandingTheme];

INSERT INTO XeroInvoiceCSV ( ContactName, EmailAddress, POAddressLine1, POCity, PORegion, POPostalCode, InvoiceNumber, Reference, InvoiceDate, DueDate, InventoryItemCode, Description, Quantity, UnitAmount, AccountCode, TaxType, BrandingTheme )
SELECT CourtDatesBTRIQ4QXero.Company AS ContactName, CourtDatesBTRIQ4QXero.EmailAddress, CourtDatesBTRIQ4QXero.Address AS POAddressLine1, CourtDatesBTRIQ4QXero.City AS POCity, CourtDatesBTRIQ4QXero.State AS PORegion, CourtDatesBTRIQ4QXero.ZIP AS POPostalCode, CourtDatesBTRIQ4QXero.InvoiceNo AS InvoiceNumber, CourtDatesID AS Reference, CourtDatesBTRIQ4QXero.InvoiceDate AS Invoicedate, CourtDatesBTRIQ4QXero.DueDate, CourtDatesBTRIQ4QXero.Code AS InventoryItemCode, ([Party1] & " v. " & [Party2] & Chr(13) & "Case Numbers:  " & [CaseNumber1] & "   |   " & [CaseNumber2] & Chr(13) & "Hearing Date:  " & [HearingDate] & Chr(13) & "Approx. " & [AudioLength] & " Minutes") AS Description, CourtDatesBTRIQ4QXero.Quantity, CourtDatesBTRIQ4QXero.[Rate] AS UnitAmount, 400 AS AccountCode, CourtDatesBTRIQ4QXero.TaxType, CourtDatesBTRIQ4QXero.BrandingThemes_BrandingTheme AS BrandingTheme
FROM CourtDatesBTRIQ4QXero
WHERE [Reference]=Forms![NewMainMenu]![ProcessJobSubformNMM].Form![JobNumberField];

UPDATE XeroInvoiceCSV SET XeroInvoiceCSV.InventoryItemCode = [CourtDatesBTRIQ4QXero].[Code], XeroInvoiceCSV.BrandingTheme = [CourtDatesBTRIQ4QXero].[BrandingTheme_BrandingTheme], XeroInvoiceCSV.UnitAmount = [CourtDatesBTRIQ4QXero].[List Price]
WHERE (([XeroInvoiceCSV].[Reference]=[CourtDatesBTRIQ4QXero].[Reference]));

SELECT Rates.ID AS RatesID, Rates.Code, Rates.[List Price], XeroInvoiceCSV.ID AS XeroInvoiceCSVID, XeroInvoiceCSV.InventoryItemCode, XeroInvoiceCSV.UnitAmount
FROM Rates INNER JOIN XeroInvoiceCSV ON XeroInvoiceCSV.[InventoryItemCode]=
Rates.[ID];

SELECT InvoicesQuery4.CourtDatesID, InvoicesQuery4.Reference, InvoicesQuery4.HearingDate, InvoicesQuery4.HearingStartTime, InvoicesQuery4.HearingEndTime, InvoicesQuery4.CasesID, InvoicesQuery4.OrderingID, InvoicesQuery4.AudioLength, InvoicesQuery4.Location, InvoicesQuery4.TurnaroundTimesCD, InvoicesQuery4.Expr1010, InvoicesQuery4.Cases_ID, InvoicesQuery4.Party1, InvoicesQuery4.Party2, InvoicesQuery4.CaseNumber1, InvoicesQuery4.CaseNumber2, InvoicesQuery4.Jurisdiction, InvoicesQuery4.CustomersID, InvoicesQuery4.Company, InvoicesQuery4.FirstName, InvoicesQuery4.LastName, InvoicesQuery4.Address, InvoicesQuery4.City, InvoicesQuery4.State, InvoicesQuery4.ZIP, InvoicesQuery4.EmailAddress, InvoicesQuery4.InvoiceNo, InvoicesQuery4.Quantity, InvoicesQuery4.InventoryItemCode, InvoicesQuery4.DueDate, InvoicesQuery4.InvoiceDate, InvoicesQuery4.AccountCode, InvoicesQuery4.TaxType, InvoicesQuery4.BrandingTheme, Rates.ID, Rates.[List Price], Rates.Code
FROM Rates INNER JOIN InvoicesQuery4 ON Rates.[ID] = InvoicesQuery4.[InventoryItemCode];

SELECT Rates.[ID], Rates.[Code], Rates.[List Price], InvoicesQuery4.[UnitAmount], InvoicesQuery4.[InventoryItemCode]
FROM Rates INNER JOIN InvoicesQuery4 ON Rates.[ID] = InvoicesQuery4.[InventoryItemCode];


