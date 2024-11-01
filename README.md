# csharp-SalesOrderEOD
This reads a specified users emails looking for an email in a date range with a specified sender and takes the most recently received email and downloads the attached .csv file. Then extracts the order confirmations from the .csv file and updates the order status in our system of matching orders to compelete while adding their consignments. An email report is sent of the orders that have been updated or unable to be updated for the support team to follow up.

This requires Azure Active Directory> Application Registraton:
 - Mail.Send
 - Mail.ReadWrite

The missing file "appsettings.json" for credentials is structured like:

{
    "GraphAPI": {
      "TenantId": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
      "ClientId": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
      "ClientSecret": "XXXXXXXXXXXXXXXXXXXXXX-XXXXXXXXXXXXXXXX",
      "ScopesUrl": "https://graph.microsoft.com/.default"
    },
    "Unleashed": {
      "ApiId": "XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX",
      "ApiKey": "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX",
      "ApiUrl": "https://api.unleashedsoftware.com/SalesOrders",
      "OrderSearchStatus":"3PL-To Pick",
      "OrderUpdateStatus":"Complete",
      "ContentType":"application/json",
      "Accept":"application/json",
      "ClientType":"Sandbox-Billson's Beverages Pty Ltd/james_tynan_order_testing"
    },
    "Email": {
        "SearchEmailSender": "XXXX@XXXX.com",
        "SearchEmailSubject": "EOD Summary Report",
        "SearchEmailInbox": "XXXX@XXXX.com.au",
        "SenderEmail": "XXXX@XXXX.com.au",
        "Recipients": [
          "XXXX@XXXX.com.au",
          "XXXX@XXXX.com.au",
          "XXXX@XXXX.com.au"
        ]
      },
    "Other": {
        "LogFileName": "program.log",
        "LogFileDirectory": ""
    }
  }