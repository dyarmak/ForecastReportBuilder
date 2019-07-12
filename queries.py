
creditQuery = """SELECT        tbl_InvoiceSummary.SubProjectID, tbl_SubProject.ProjectManager, tbl_InvoiceSummary.InvoiceCompanyName, tbl_MasterProject.MasterProjectName, tbl_Projects.ProjectID, tbl_Projects.ProjectName, 
                         tbl_SubProjectType_2.SubProjectTypeName, tbl_SubProject.SubProjectType AS SubProjectTypeID, tbl_SubProject.SubProjectName, tbl_InvoiceSummary.SubTotal - tbl_InvoiceSummary.OriginalSubTotal AS CreditAmt, 
                         tbl_InvoiceSummary.CreditDateApplied AS [Due Date], tbl_SubProject.SubProjectStatus, tbl_SubProject.Quoted, tbl_SubProject.Forecast AS OriginalForecast, tbl_SubProject.DueDate AS OriginalDueDate
FROM            tbl_MasterProject INNER JOIN
                         tbl_Projects ON tbl_MasterProject.MasterProjectID = tbl_Projects.MasterProjectID INNER JOIN
                         tbl_InvoiceSummary INNER JOIN
                         tbl_SubProject ON tbl_InvoiceSummary.SubProjectID = tbl_SubProject.SubProjectID ON tbl_Projects.ProjectID = tbl_SubProject.ProjectID INNER JOIN
                         tbl_SubProjectType AS tbl_SubProjectType_2 ON tbl_SubProject.SubProjectType = tbl_SubProjectType_2.SubProjectTypeID
WHERE        (tbl_InvoiceSummary.CreditDateApplied BETWEEN '1/1/2019' AND '12/31/2019')
ORDER BY tbl_InvoiceSummary.SubProjectID;"""

invoicedQuery = """SELECT        tbl_InvoiceSummary.SubProjectID, tbl_SubProject.ProjectManager, tbl_InvoiceSummary.InvoiceCompanyName, tbl_MasterProject.MasterProjectName, tbl_Projects.ProjectID, tbl_Projects.ProjectName, 
                         tbl_SubProjectType_2.SubProjectTypeName, tbl_SubProject.SubProjectType AS SubProjectTypeID, tbl_SubProject.SubProjectName, tbl_InvoiceSummary.OriginalSubTotal, 
                         tbl_InvoiceSummary.RecognizeRevenueDate AS [Due Date], tbl_SubProject.SubProjectStatus, tbl_SubProject.Quoted, tbl_SubProject.Forecast AS OriginalForecast, tbl_SubProject.DueDate AS OriginalDueDate, 
                         tbl_InvoiceSummary.InvoiceID, tbl_InvoiceSummary.InvoiceDateSent
FROM            tbl_MasterProject INNER JOIN
                         tbl_Projects ON tbl_MasterProject.MasterProjectID = tbl_Projects.MasterProjectID INNER JOIN
                         tbl_InvoiceSummary INNER JOIN
                         tbl_SubProject ON tbl_InvoiceSummary.SubProjectID = tbl_SubProject.SubProjectID ON tbl_Projects.ProjectID = tbl_SubProject.ProjectID INNER JOIN
                         tbl_SubProjectType AS tbl_SubProjectType_2 ON tbl_SubProject.SubProjectType = tbl_SubProjectType_2.SubProjectTypeID
WHERE        (tbl_InvoiceSummary.RecognizeRevenueDate BETWEEN '1/1/2019' AND '12/31/2019')
ORDER BY tbl_InvoiceSummary.SubProjectID;"""

forecastQuery = """SELECT        tbl_SubProject.SubProjectID, tbl_SubProject.ProjectManager, tbl_Client.ClientName, tbl_MasterProject.MasterProjectName, tbl_SubProject.ProjectID, tbl_Projects.ProjectName, tbl_SubProjectType.SubProjectTypeName, 
                         tbl_SubProject.SubProjectType AS SubProjectTypeID, tbl_SubProject.SubProjectName, tbl_SubProject.Forecast, tbl_SubProject.DueDate + 10 AS [Due Date], tbl_SubProject.SubProjectStatus, tbl_SubProject.Quoted, 
                         tbl_SubProject.Forecast AS OriginalForecast, tbl_SubProject.DueDate AS OriginalDueDate, tbl_SubProject.Budget, tbl_SubProject.ClientBudgetYear, tbl_SubProject.IncludeinBudget, tbl_InvoiceSummary.InvoiceID, 
                         tbl_InvoiceSummary.InvoiceDateSent, tbl_InvoiceSummary.SubTotal
FROM            tbl_SubProjectType INNER JOIN
                         tbl_Client INNER JOIN
                         tbl_MasterProject INNER JOIN
                         tbl_Projects ON tbl_MasterProject.MasterProjectID = tbl_Projects.MasterProjectID ON tbl_Client.ClientID = tbl_MasterProject.ClientID INNER JOIN
                         tbl_SubProject ON tbl_Projects.ProjectID = tbl_SubProject.ProjectID ON tbl_SubProjectType.SubProjectTypeID = tbl_SubProject.SubProjectType LEFT OUTER JOIN
                         tbl_InvoiceSummary ON tbl_SubProject.SubProjectID = tbl_InvoiceSummary.SubProjectID
WHERE        (tbl_Client.ClientName <> 'H2Safety') AND (tbl_SubProject.SubProjectStatus <> 'Invoiced') AND (tbl_SubProject.SubProjectStatus <> 'Cancelled')
ORDER BY tbl_SubProject.SubProjectID;"""

vlookQuery = """SELECT        tbl_SubProjectType.SubProjectTypeName, tbl_SubProjectCategories.CategoryDescription AS ProfitCenter, tbl_SubProjectType.SubProjectTypeID
FROM            tbl_SubProjectType INNER JOIN
                         tbl_SubProjectTypeCategories ON tbl_SubProjectType.SubProjectTypeID = tbl_SubProjectTypeCategories.SubProjectTypeID INNER JOIN
                         tbl_SubProjectCategories ON tbl_SubProjectTypeCategories.SubProjectCategoryID = tbl_SubProjectCategories.SubProjectCategoryID
ORDER BY tbl_SubProjectTypeCategories.SubProjectTypeCategoryID;"""
