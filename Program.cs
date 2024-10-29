using Azure.Identity;
using Azure.Storage.Blobs;
using CsvHelper;
using CsvHelper.Configuration;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Me.SendMail;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions.Extensions;
using Newtonsoft.Json;
using RestSharp;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace csharpConsolegraphtest
{
    class Program
    {
        static async Task Main(string[] args)
        {
            // Set up for logging and loading configurations
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory()) // Assuming the file is in the same folder as the executable
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            Configuration = builder.Build();
            string logFileName = Configuration["Other:LogFileName"];

            // Step 0. Lets setup program logs sent to a file.
            string logFilePath = Path.Combine(Environment.CurrentDirectory, logFileName);
            using (StreamWriter logFile = new StreamWriter(logFilePath, append: false))
            {
                logFile.AutoFlush = true;
                Console.SetOut(logFile);
                Console.WriteLine("Program Started...");

                // Step 1. Set up the configuration to read the appsettings.json file then read in the config
                var graphApiConfig = new GraphAPIConfig();
                Configuration.GetSection("GraphAPI").Bind(graphApiConfig);
                ValidateConfig(graphApiConfig);

                var unleashedConfig = new UnleashedConfig();
                Configuration.GetSection("Unleashed").Bind(unleashedConfig);
                ValidateConfig(unleashedConfig);

                var emailConfig = new EmailConfig();
                Configuration.GetSection("Email").Bind(emailConfig);
                ValidateConfig(emailConfig);

                // Create a credential object using ClientSecretCredential
                var clientSecretCredential = new ClientSecretCredential(graphApiConfig.TenantId, graphApiConfig.ClientId, graphApiConfig.ClientSecret);
                
                // Define the required scopes for the Graph API. For reading emails, we need Mail.Read.
                var scopes = new[] { graphApiConfig.ScopesUrl };
                
                // Initialise GraphServiceClient with the credential.
                var graphClient = new GraphServiceClient(clientSecretCredential, scopes);

                // Step 2. Search sales emails and download the CSV file from an email attachment.
                string csvFilePath = await FindAndDownloadCsvEmailAsync(graphClient, emailConfig);

                // Step 3. If a CSV file was downloaded, parse it
                if (!string.IsNullOrEmpty(csvFilePath))
                {
                    var salesOrders = await ParseCsvFileAsync(csvFilePath);
                    var salesOrdersUpload = salesOrders
                        .Select(csvOrder => new CsvRecordUpload(csvOrder))
                        .ToList();

                    // Step 4. Retrieve "3pl-to-pick" orders from Unleashed
                    var unleashedOrders = await GetSalesOrdersWithStatus(unleashedConfig);

                    if (unleashedOrders == null)
                    {
                        Console.WriteLine("Failed to retrieve Unleashed orders.");
                        return;
                    }

                    // Step 5. Cross-reference CSV orders with Unleashed orders and update status
                    foreach (var salesOrderUpload in salesOrdersUpload)
                    {
                        var matchingOrder = unleashedOrders.FirstOrDefault(
                            uOrder => uOrder.OrderNumber == salesOrderUpload.SalesOrder);

                        if (matchingOrder != null)
                        {
                            // If a match is found, update the order information and status
                            salesOrderUpload.OrderDate = matchingOrder.OrderDate;
                            salesOrderUpload.OrderTotal = matchingOrder.SalesOrderLines.Sum(line => line.LineTotal);
                            salesOrderUpload.UploadStatus = "Found in Unleashed.";

                            // Step 5. Update order status using the matching order
                            bool success = await UpdateSalesOrderStatusAsync(matchingOrder, unleashedConfig, salesOrderUpload.ConsignmentNumber);
                            if (success)
                            {
                                salesOrderUpload.UploadStatus = "Successfully updated order in Unleashed.";
                                Console.WriteLine($"SalesOrder {salesOrderUpload.SalesOrder} successfully updated to '{unleashedConfig.OrderUpdateStatus}'.");
                            }
                            else
                            {
                                salesOrderUpload.UploadStatus = "Failed to update order in Unleashed.";
                                Console.WriteLine($"Failed to update SalesOrder {salesOrderUpload.SalesOrder}.");
                            }

                        }
                        else 
                        {
                            salesOrderUpload.UploadStatus = "Failed to find order in Unleashed.";
                            Console.WriteLine($"No matching '{unleashedConfig.OrderSearchStatus}' Unleashed order found for CSV order: {salesOrderUpload.SalesOrder}");
                        }
                    }

                    // Step 7. Recheck the order statuses and send email report (implement email logic here)
                    // Example call for sending email (needs actual email logic):
                    await SendStatusReportEmail(graphClient, salesOrdersUpload, emailConfig);
                }
                else
                {
                    Console.WriteLine("No CSV file was downloaded.");
                }
                Console.WriteLine("Program completed...");
                logFile.Close();
            }       
        }

        static async Task<string> FindAndDownloadCsvEmailAsync(GraphServiceClient graphClient, EmailConfig emailConfig)
        {

            try
            {
                // Get the current date in UTC, then format it.
                DateTime todayLocal = DateTime.UtcNow.ToLocalTime();
                todayLocal = todayLocal.Date.AddDays(0); // strip out the time component and get start of day (midnight).
                string todayStartLocal = todayLocal.ToString("yyyy-MM-ddTHH:mm:ssZ");
                Console.WriteLine($"Today Start Local: {todayStartLocal}");
                
                // Build our filter parameters, Subject, email address, date and time.
                string queryParametersFilter = $"receivedDateTime ge {todayStartLocal}";                
                if (!string.IsNullOrEmpty(emailConfig.SearchEmailSender))
                {
                    queryParametersFilter += $" and from/emailAddress/address eq '{emailConfig.SearchEmailSender}'";
                }
                if (!string.IsNullOrEmpty(emailConfig.SearchEmailSubject))
                {
                    queryParametersFilter += $" and contains(subject, '{emailConfig.SearchEmailSubject}')";
                }

                // Get the signed-in user's email messages.
                var messages = await graphClient.Users[emailConfig.SearchEmailInbox]
                    .Messages
                    .GetAsync(requestConfiguration =>
                    {
                        requestConfiguration.QueryParameters.Filter = queryParametersFilter;
                        requestConfiguration.QueryParameters.Top = 1;
                    });

                // Print out the subject of each email.
                if (messages?.Value?.Count > 0)
                {
                    foreach (var message in messages.Value)
                    {
                        Console.WriteLine($"Found message: {message.Subject}");

                        // Get the attachments for the message
                        var attachments = await graphClient.Users[emailConfig.SearchEmailInbox]
                            .Messages[message.Id]
                            .Attachments
                            .GetAsync();

                        // Check if the message has any attachments
                        if (attachments?.Value?.Count > 0)
                        {
                            foreach (var attachment in attachments.Value)
                            {
                                if (attachment is FileAttachment fileAttachment)
                                {
                                    // Check is the attached file is a .csv file
                                    if (fileAttachment.Name.EndsWith(".csv", StringComparison.OrdinalIgnoreCase))
                                    {
                                        Console.WriteLine($"Downloading CSV attachment: {fileAttachment.Name}");
                                                                        // Path to save the attachment
                                        var filePath = Path.Combine(Environment.CurrentDirectory, fileAttachment.Name);

                                        // Write the file content to the local file
                                        await File.WriteAllBytesAsync(filePath, fileAttachment.ContentBytes);

                                        Console.WriteLine($"Attachment saved to: {filePath}");
                                        return filePath; 
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Skipping non-CSV attachment {fileAttachment.Name}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No attachments found in the email.");
                        }
                    }
                }
                else
                {
                    Console.WriteLine("No messages found with the specificied subject and date.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving emails: {ex.Message}");
            }
            return null;
        }

        static async Task<List<CsvRecord>> ParseCsvFileAsync(string csvFilePath)
        {
            var records = new List<CsvRecord>();

            // Open the file for reading
            using (var reader = new StreamReader(csvFilePath))
            using (var csv = new CsvReader(reader, CultureInfo.InvariantCulture))
            {
                // Read the file and map it to CsvRecord objects
                await foreach (var record in csv.GetRecordsAsync<CsvRecord>())
                {
                    records.Add(record);
                }

                // Example of accessing data from the CsvRecord list
                foreach (var record in records)
                {
                    Console.WriteLine($"SalesOrder: {record.SalesOrder}, Name: {record.ShipToContact}, Address Line 1: {record.ShipToAdd1}, ConsignmentNumber: {record.ConsignmentNumber}");
                }
            }
            return records;
        }

        static async Task<List<SalesOrder>> GetSalesOrdersWithStatus(UnleashedConfig unleashedConfig)
        {
            // API endpoint for retrieving sales with the given status
            string apiQueryString = $"CustomOrderStatus={unleashedConfig.OrderSearchStatus}";

            // Generate the signature using the query string
            string signature = GenerateSignature(apiQueryString, unleashedConfig.ApiKey);

            // Set up the request
            var client = new RestClient($"{unleashedConfig.ApiUrl}?{apiQueryString}");
            var request = new RestRequest();
            request.Method = Method.Get;

            // Add the required headers
            request.AddHeader("Content-Type", unleashedConfig.ContentType);
            request.AddHeader("Accept", unleashedConfig.Accept);
            request.AddHeader("api-auth-id", unleashedConfig.ApiId);
            request.AddHeader("api-auth-signature", signature);
            request.AddHeader("client-type", unleashedConfig.ClientType);

            try
            {
                // Execute the request and get the response
                RestResponse response = await client.ExecuteAsync(request);

                if (response.IsSuccessful)
                {
                    // Deserialize the response into the Root object
                    var settings = new JsonSerializerSettings
                    {
                        DateFormatHandling = DateFormatHandling.MicrosoftDateFormat
                    };
                    
                    var result = JsonConvert.DeserializeObject<Root>(response.Content, settings);

                    // Return the list of SalesOrders (Items)
                    return result.Items;
                }
                else
                {
                    Console.WriteLine($"Failed to retrieve sales orders. Status code: {response.StatusCode}");
                    return null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving sales orders: {ex.Message}");
                return null;
            }
        }

        static async Task<bool> UpdateSalesOrderStatusAsync(SalesOrder existingOrder, UnleashedConfig unleashedConfig, string consignmentNumber)
        {
            // Prepare the URL using the order GUID
            string apiUrlFull = $"{unleashedConfig.ApiUrl}/{existingOrder.Guid}";
            string apiQuery = "";

            // Create a reduced version of the SalesOrder object
            var uploadSalesOrder = PrepareSalesOrderJson(existingOrder, unleashedConfig.OrderUpdateStatus,  consignmentNumber);

            // Serialize the reduced sales order into JSON
            string jsonPayload = JsonConvert.SerializeObject(uploadSalesOrder);
            string jsonPayloadTesting = JsonConvert.SerializeObject(uploadSalesOrder, Formatting.None);
            // Console.WriteLine(JsonConvert.SerializeObject(uploadSalesOrder, Formatting.Indented));

            // Generate the signature using our blank query and Api Key
            string signature = GenerateSignature(apiQuery, unleashedConfig.ApiKey);

            // Create the RestClient and RestRequest
            var client = new RestClient(apiUrlFull);
            var request = new RestRequest();
            request.Method = Method.Put;

            // Add the required headers
            request.AddHeader("Content-Type", unleashedConfig.ContentType);
            request.AddHeader("Accept", unleashedConfig.Accept);
            request.AddHeader("api-auth-id", unleashedConfig.ApiId);
            request.AddHeader("api-auth-signature", signature);
            request.AddHeader("client-type", unleashedConfig.ClientType);

            // Add the JSON payload to the request
            request.AddJsonBody(uploadSalesOrder);

            try
            {
                // Execute the request and get the response
                RestResponse response = await client.ExecuteAsync(request);

                if (response.IsSuccessful)
                {
                    Console.WriteLine($"Successfully updated SalesOrder {existingOrder.OrderNumber} to '{unleashedConfig.OrderUpdateStatus}'.");
                    return true;
                }
                else
                {
                    Console.WriteLine($"Failed to update SalesOrder {existingOrder.OrderNumber}. Status code: {response.StatusCode}");
                    Console.WriteLine($"API URL : {apiUrlFull}");
                    return false;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating SalesOrder {existingOrder.OrderNumber}: {ex.Message}");
                return false;
            }
        }

        public static string PrepareMinimalSalesOrderJson(SalesOrder existingOrder, string newStatus, string consignmentNumber)
        {
            // Prepare a minimal object with only the fields that should be updated
            var minimalSalesOrder = new
            {
                Guid = existingOrder.Guid,  // Unleashed requires the GUID for identification
                OrderStatus = newStatus,    // Update order status to "Completed" or another value
                Comments = $"Consignment Number: {consignmentNumber}.", // Update comments with consignment info
                SalesOrderLines = existingOrder.SalesOrderLines.Select(line => new
                {
                    line.LineNumber,  // You must include at least one sales order line to keep it valid
                    line.Guid  // Include the GUID for each sales order line
                }).ToList()
            };

            // Serialize the object to JSON using Newtonsoft.Json
            string jsonPayload = JsonConvert.SerializeObject(minimalSalesOrder, Formatting.Indented);
            return jsonPayload;
        }

        public static string PrepareSalesOrderJson(SalesOrder existingOrder, string newStatus, string consignmentNumber)
        {
            // Prepare the reduced sales order with the necessary fields for the PUT request
            var updatedSalesOrder = new
            {
                Comments = $"{existingOrder.Comments}  Consignment Number: {consignmentNumber}",  // Add consignment number to comments
                CustomerRef = existingOrder.CustomerRef,
                DeliveryCity = existingOrder.DeliveryCity,
                DeliveryCountry = existingOrder.DeliveryCountry,
                DeliveryInstruction = existingOrder.DeliveryInstruction,
                DeliveryMethod = existingOrder.DeliveryMethod,
                DeliveryName = existingOrder.DeliveryName,
                DeliveryPostCode = existingOrder.DeliveryPostCode,
                DeliveryRegion = existingOrder.DeliveryRegion,
                DeliveryStreetAddress = existingOrder.DeliveryStreetAddress,
                DeliveryStreetAddress2 = existingOrder.DeliveryStreetAddress2,
                DeliverySuburb = existingOrder.DeliveryCountry,
                DiscountRate = existingOrder.ExchangeRate,
                ExchangeRate = existingOrder.ExchangeRate,
                //OrderDate = ParseDate(existingOrder.OrderDate),  // Convert date if needed
                //OrderNumber = existingOrder.OrderNumber,
                OrderStatus = newStatus,  // Update order status to the new status
                //ReceivedDate = existingOrder.ReceivedDate,
                RequiredDate = ConvertToIso8601(existingOrder.RequiredDate),  // Convert date if needed
                SalesOrderGroup = existingOrder.SalesOrderGroup,
                SalesOrderLines = existingOrder.SalesOrderLines.Select(line => new
                {
                    line.LineNumber,
                    Product = new
                    {
                        line.Product.Guid,
                        line.Product.ProductCode
                    },
                    line.OrderQuantity,
                    line.UnitPrice,
                    line.LineTotal,
                    line.TaxRate,
                    line.LineTax,
                    line.XeroTaxCode,
                    line.Guid,
                    SerialNumbers = line.SerialNumbers?.Select(sn => new
                    {
                        sn.Identifier
                    }),
                    BatchNumbers = line.BatchNumbers?.Select(bn => new
                    {
                        bn.Number,
                        bn.Quantity
                    })
                }).ToList(),
                Salesperson = new
                {
                    FullName = existingOrder.Salesperson.FullName,
                    Email = existingOrder.Salesperson.Email,
                    Obsolete = existingOrder.Salesperson.Obsolete,
                    Guid = existingOrder.Salesperson.Guid,
                    SourceId = existingOrder.Salesperson.LastModifiedOn
                },
                SourceId = existingOrder.SourceId,
                //SubTotal = existingOrder.SubTotal,
                Tax = new
                {
                    existingOrder.Tax.TaxCode,
                    existingOrder.Tax.TaxRate
                },
                //TaxRate = existingOrder.TaxRate,
                //TaxTotal = existingOrder.TaxTotal,
                Warehouse = new
                {
                    existingOrder.Warehouse.WarehouseCode,
                    existingOrder.Warehouse.Guid
                }
            };

            // Serialize the object to JSON using Newtonsoft.Json
            string jsonPayload = JsonConvert.SerializeObject(updatedSalesOrder, Formatting.Indented);
            return jsonPayload;
        }

        // Function to convert /Date(UnixTimestamp)/ format to ISO 8601 (yyyy-MM-ddTHH:mm:ssZ)
        static string ConvertToIso8601(string jsonDate)
        {
            // Remove the '/Date(' prefix and ')/' suffix
            if (jsonDate.StartsWith("/Date(") && jsonDate.EndsWith(")/"))
            {
                string timestamp = jsonDate.Substring(6, jsonDate.Length - 8);  // Extract the timestamp
                if (long.TryParse(timestamp, out long milliseconds))
                {
                    // Convert milliseconds to DateTime
                    DateTime dateTime = DateTimeOffset.FromUnixTimeMilliseconds(milliseconds).UtcDateTime;
                    // Return in ISO 8601 format (adjust format as necessary)
                    return dateTime.ToString("yyyy-MM-ddTHH:mm:ssZ");
                }
            }
            return jsonDate;  // Return the original value if conversion fails
        }

        static async Task SendEmailAsync(GraphServiceClient graphClient,  string senderEmail, List<string> recipientEmails, string subject, string bodyContent)
        {
            var message = new Message
            {
                Subject = subject,
                Body = new ItemBody
                {
                    ContentType = BodyType.Html, // Using HTML for better formatting
                    Content = bodyContent
                },
                ToRecipients = new List<Recipient>() // Initialize ToRecipients list
            };

            // Add each recipient email to the ToRecipients list
            foreach (var recipientEmail in recipientEmails)
            {
                message.ToRecipients.Add(new Recipient
                {
                    EmailAddress = new EmailAddress
                    {
                        Address = recipientEmail
                    }
                });
            }

            var sendMailBody = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody
            {
                Message = message,
                SaveToSentItems = true // Save email in Sent Items
            };

            try
            {
                // Send the email on behalf of the authenticated user
                await graphClient.Users[senderEmail]
                    .SendMail
                    .PostAsync(sendMailBody);

                Console.WriteLine("Email sent successfully to all recipients.");
            }
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }


        static async Task SendStatusReportEmail(GraphServiceClient graphClient, List<CsvRecordUpload> csvOrders, EmailConfig emailConfig)
        {
            DateTime todayLocal = DateTime.UtcNow.ToLocalTime();
            string todayStartLocal = todayLocal.ToString("yyyy-MM-dd");
            string emailHeader = $"Order Status Update {todayStartLocal}";

            // Build the HTML content for the email
            StringBuilder reportContent = new StringBuilder();
            reportContent.AppendLine("<h2>Order Status Upload Report</h2>");
            reportContent.AppendLine("<p>Here is the upload report on the 3PL-To Pick orders:</p>");
            
            // Create a table for the order details
            reportContent.AppendLine("<table style='border-collapse: collapse; width: 100%;'>");
            reportContent.AppendLine("<thead>");
            reportContent.AppendLine("<tr style='background-color: #f2f2f2;'>");
            reportContent.AppendLine("<th style='border: 1px solid #ddd; padding: 8px;'>Order</th>");
            reportContent.AppendLine("<th style='border: 1px solid #ddd; padding: 8px;'>Order Status</th>");
            reportContent.AppendLine("<th style='border: 1px solid #ddd; padding: 8px;'>Customer</th>");
            reportContent.AppendLine("<th style='border: 1px solid #ddd; padding: 8px;'>Order Consignment</th>");
            reportContent.AppendLine("</tr>");
            reportContent.AppendLine("</thead>");
            reportContent.AppendLine("<tbody>");
                
            // Add rows for each order
            foreach (var order in csvOrders)
            {
                reportContent.AppendLine("<tr>");
                reportContent.AppendLine($"<td style='border: 1px solid #ddd; padding: 8px;'>{order.SalesOrder}</td>");
                reportContent.AppendLine($"<td style='border: 1px solid #ddd; padding: 8px;'>{order.UploadStatus}</td>");
                reportContent.AppendLine($"<td style='border: 1px solid #ddd; padding: 8px;'>{order.ShipToContact}</td>");
                reportContent.AppendLine($"<td style='border: 1px solid #ddd; padding: 8px;'>{order.ConsignmentNumber}</td>");
                reportContent.AppendLine("</tr>");
            }

            reportContent.AppendLine("</tbody>");
            reportContent.AppendLine("</table>");
            reportContent.AppendLine("<p style='font-size: 12px; color: #888;'>This is an automated report from the order management system.</p>");

            await SendEmailAsync(graphClient, emailConfig.SenderEmail, emailConfig.Recipients, emailHeader, reportContent.ToString());
        }

        // Generate HMAC-SHA256 signature using only the query string
        static string GenerateSignature(string queryString, string apiKey)
        {
            if (string.IsNullOrEmpty(queryString))
            {
                queryString = string.Empty;  // Handle empty query strings
            }

            // Convert the API key and query string to byte arrays
            var encoding = new UTF8Encoding();
            byte[] keyBytes = encoding.GetBytes(apiKey);
            byte[] messageBytes = encoding.GetBytes(queryString);

            // Create HMACSHA256 hash using the API key
            using (var hmacsha256 = new HMACSHA256(keyBytes))
            {
                byte[] hashValue = hmacsha256.ComputeHash(messageBytes);
                return Convert.ToBase64String(hashValue);  // Convert the hash to Base64
            }
        }

        public class CsvRecordUpload: CsvRecord
        {
            public string? UploadStatus { get; set; } = "Not order found matching in Unleashed.";
            public string? OrderDate { get; set; }
            public decimal? OrderTotal { get; set; }

            public CsvRecordUpload(CsvRecord csvRecord)
            {
                // Copy data from the base CsvRecord class
                SalesOrder = csvRecord.SalesOrder;

                CustomerOrderNumber = csvRecord.CustomerOrderNumber;
                ShipToContact = csvRecord.ShipToContact;
                ShipToAdd1 = csvRecord.ShipToAdd1;
                ShipToAdd2 = csvRecord.ShipToAdd2;
                ShipToAdd3 = csvRecord.ShipToAdd3;
                ShipToSuburb = csvRecord.ShipToSuburb;
                ShipToState = csvRecord.ShipToState;
                ShipToPostCode = csvRecord.ShipToPostCode;
                ShipToCountry = csvRecord.ShipToCountry;
                CreatedDate = csvRecord.CreatedDate;
                PickedDate = csvRecord.PickedDate;
                DispatchedDate = csvRecord.DispatchedDate;
                ConsignmentNumber = csvRecord.ConsignmentNumber;                
            }
        }

        public class CsvRecord
        {
            public string? SalesOrder { get; init; } // Unleashed sales order number
            public string? CustomerOrderNumber { get; init; }
            public string? ShipToContact { get; init; }
            public string? ShipToAdd1 { get; init; } // Shipping Address Line 1 has most of the info
            public string? ShipToAdd2 { get; init; }
            public string? ShipToAdd3 { get; init; }
            public string? ShipToSuburb { get; init; }
            public string? ShipToState { get; init; }
            public string? ShipToPostCode { get; init; }
            public string? ShipToCountry { get; init; }
            public string? CreatedDate { get; init; }
            public string? PickedDate { get; init; }
            public string? DispatchedDate { get; init; }
            public string? ConsignmentNumber { get; init; }
        }

        public class Product
        {
            public string? Guid { get; set; }
            public string ProductCode { get; set; }
            public string ProductDescription { get; set; }
        }

        public class SalesOrderLine
        {
            public int LineNumber { get; set; }
            public string? LineType { get; set; }
            public Product Product { get; set; }
            public string? DueDate { get; set; }  // Date in "/Date(...)/" format
            public decimal OrderQuantity { get; set; }
            public decimal UnitPrice { get; set; }
            public decimal DiscountRate { get; set; }
            public decimal LineTotal { get; set; }
            public string? Comments { get; set; }
            public decimal TaxRate { get; set; }
            public decimal LineTax { get; set; }
            public string? XeroTaxCode { get; set; }
            public string? Guid { get; set; }
            public string? LastModifiedOn { get; set; }  // Date in "/Date(...)/" format
            public List<SerialNumber>? SerialNumbers { get; set; }
            public List<BatchNumber>? BatchNumbers { get; set; }
        }

        public class SerialNumber
        {
            public string Identifier { get; set; }
        }

        public class BatchNumber
        {
            public string Number { get; set; }
            public string Quantity { get; set; }
        }

        public class Assembly
        {
            public string Guid { get; set; }
            public string AssemblyNumber { get; set; }
            public string AssemblyStatus { get; set; }
        }

        public class Customer
        {
            public string CustomerCode { get; set; }
            public string CustomerName { get; set; }
            public int CurrencyId { get; set; }
            public string? Guid { get; set; }
            public string? LastModifiedOn { get; set; }  // Date in "/Date(...)/" format
        }

        public class Warehouse
        {
            public string? WarehouseCode { get; set; }
            public string? WarehouseName { get; set; }
            public string? City { get; set; }
            public string? Country { get; set; }
            public string? StreetNo { get; set; }
            public string? AddressLine1 { get; set; }
            public string? AddressLine2 { get; set; }
            public string? Suburb { get; set; }
            public string? Region { get; set; }
            public string? PostCode { get; set; }
            public string? PhoneNumber { get; set; }
            public string? MobileNumber { get; set; }
            public string? ContactName { get; set; }
            public string? Guid { get; set; }
        }

        public class DeliveryContact
        {
            public string? Guid { get; set; }
            public string FirstName { get; set; }
            public string LastName { get; set; }
        }

        public class Currency
        {
            public string CurrencyCode { get; set; }
            public string? Description { get; set; }
            public string? Guid { get; set; }
            public string? LastModifiedOn { get; set; }  // Date in "/Date(...)/" format
        }

        public class Tax
        {
            public string? TaxCode { get; set; }
            public string? Description { get; set; }
            public decimal TaxRate { get; set; }
            public bool CanApplyToExpenses { get; set; }
            public bool CanApplyToRevenue { get; set; }
            public bool Obsolete { get; set; }
            public string? Guid { get; set; }
            public string? LastModifiedOn { get; set; }
        }
        public class Salesperson
        {
            public string? FullName { get; set; }
            public string? Email { get; set; }
            public bool? Obsolete { get; set; }
            public string Guid { get; set; }
            public string LastModifiedOn { get; set; }
        }

        public class SalesOrder
        {
            public string? Guid { get; set; }
            public string OrderNumber { get; set; }
            public string? OrderDate { get; set; }  // Date in "/Date(...)/" format
            public string? RequiredDate { get; set; }
            public string? CompletedDate { get; set; }  // Nullable date
            public string? OrderStatus { get; set; }
            public Customer Customer { get; set; }
            public Warehouse? Warehouse { get; set; }
            public List<SalesOrderLine>? SalesOrderLines { get; set; }
            public string? Comments { get; set; }
            public string? CustomOrderStatus { get; set; }
            public string? CustomerRef { get; set; }
            public DeliveryContact? DeliveryContact { get; set; }
            public string? DeliveryInstruction { get; set; }
            public string? DeliveryMethod { get; set; }
            public string? DeliveryName { get; set; }
            public string? DeliveryCountry { get; set; }
            public string? DeliveryCity { get; set; }
            public string? DeliverySuburb { get; set; }
            public string? DeliveryRegion { get; set; }
            public string? DeliveryPostCode { get; set; }
            public string? DeliveryStreetAddress { get; set; }
            public string? DeliveryStreetAddress2 { get; set; }
            public string? ReceivedDate { get; set; }
            public Currency? Currency { get; set; }
            public decimal ExchangeRate { get; set; }
            public decimal SubTotal { get; set; }
            public decimal TaxRate { get; set; }
            public decimal TaxTotal { get; set; }
            public decimal Total { get; set; }
            public decimal? TotalVolume { get; set; }
            public decimal? TotalWeight { get; set; }
            public decimal DiscountRate { get; set; }
            public Tax? Tax { get; set; }
            public string? XeroTaxCode { get; set; }
            public string? PaymentDueDate { get; set; }  // Nullable date
            public bool AllocateProduct { get; set; }
            public string? SalesOrderGroup { get; set; }
            public Salesperson? Salesperson { get; set; }
            public bool SendAccountingJournalOnly { get; set; }
            public string? SourceId { get; set; }
            public string? CreatedBy { get; set; }
            public string? CreatedOn { get; set; }  // Date in "/Date(...)/" format
            public string? LastModifiedBy { get; set; }
            public string? LastModifiedOn { get; set; }  // Date in "/Date(...)/" format
        }

        public class Pagination
        {
            public int NumberOfItems { get; set; }
            public int PageSize { get; set; }
            public int PageNumber { get; set; }
            public int NumberOfPages { get; set; }
        }

        public class Root
        {
            public Pagination Pagination { get; set; }
            public List<SalesOrder> Items { get; set; }
        }

        public class GraphAPIConfig : IValidatable
        {
            public string TenantId { get; set; }
            public string ClientId { get; set; }
            public string ClientSecret { get; set; }
            public string ScopesUrl { get; set; }
            public bool IsValid(out List<string> errors)
            {
                errors = new List<string>();

                if (string.IsNullOrEmpty(TenantId))
                {
                    errors.Add("The Graph TenantId is required. (appsettings.json>GraphAPI>TenantId)");
                }
                if (string.IsNullOrEmpty(ClientId))
                {
                    errors.Add("The Graph ClientId is required. (appsettings.json>GraphAPI>CleintId)");
                }
                if (string.IsNullOrEmpty(ClientSecret))
                {
                    errors.Add("The Graph client secret is required. (appsettings.json>GraphAPI>ClientSecret)");
                }
                if (string.IsNullOrEmpty(ScopesUrl))
                {
                    errors.Add("The Graph scopes URL is required. (appsettings.json>GraphAPI>ScopesUrl)");
                }
                // we use the number of errors to return a boolean, 0 = IsValid returns true
                return errors.Count == 0;
            }            
        }

        public class UnleashedConfig : IValidatable
        {
            public string ApiId { get; set; }
            public string ApiKey { get; set; }
            public string ApiUrl { get; set; }
            public string OrderSearchStatus { get; set; }
            public string OrderUpdateStatus { get; set; }
            public string ContentType { get; set; }
            public string Accept { get; set; }
            public string ClientType { get; set; }
            public bool IsValid(out List<string> errors)
            {
                errors = new List<string>();

                if (string.IsNullOrEmpty(ApiId))
                {
                    errors.Add("Unleashed API ID is required. (appsettings.json>Unleashed>ApiId)");
                }
                if (string.IsNullOrEmpty(ApiKey))
                {
                    errors.Add("Unleashed API Key is required. (appsettings.json>Unleashed>ApiKey)");
                }
                if (string.IsNullOrEmpty(ApiUrl))
                {
                    errors.Add("Unleashed API Url is required. (appsettings.json>Unleashed>ApiUrl)");
                }
                if (string.IsNullOrEmpty(OrderSearchStatus))
                {
                    errors.Add("The order status to filter our returns is required. (appsettings.json>Unleashed>OrderSearchStatus)");
                }
                if (string.IsNullOrEmpty(OrderUpdateStatus))
                {
                    errors.Add("The order status to update Unleashed orders with is required. (appsettings.json>Unleashed>OrderUpdateStatus)");
                }
                if (string.IsNullOrEmpty(ContentType))
                {
                    errors.Add("The returned data type is required. (appsettings.json>Unleashed>ContentType)");
                }
                if (string.IsNullOrEmpty(Accept))
                {
                    errors.Add("The accepted data type is required. (appsettings.json>Unleashed>Accept)");
                }
                if (string.IsNullOrEmpty(ClientType))
                {
                    errors.Add("The business and developer 'signature' is required. (appsettings.json>Unleashed>ClientType)");
                }
                // we use the number of errors to return a boolean, 0 = IsValid returns true
                return errors.Count == 0;
            }
        }

        public class EmailConfig : IValidatable
        {
            public string ?SearchEmailSender { get; set; }
            public string ?SearchEmailSubject { get; set; }
            public string SearchEmailInbox { get; set; }
            public string SenderEmail { get; set; }
            public List<string> Recipients { get; set; }

            public bool IsValid(out List<string> errors)
            {
                errors = new List<string>();

                if (string.IsNullOrEmpty(SenderEmail))
                {
                    errors.Add("Email address to send as is required. (appsettings.json>Email>SenderEmail)");
                }
                if (!(Recipients.Count > 0))
                {
                    errors.Add("We require at least one recipient email address. (appsettings.json>Email>Recipients{})");
                }
                // we use the number of errors to return a boolean, 0 = IsValid returns true
                return errors.Count == 0;
            }
        }

        public static IConfigurationRoot Configuration;
        public interface IValidatable
        {
            bool IsValid(out List<string> errors);
        }

        public static void ValidateConfig(IValidatable config)
        {
            if (!config.IsValid(out List<string> errors))
            {
                throw new ConfigurationException($"{config.GetType().Name} has configuration errors:\n{string.Join("\n", errors)}");
            }
        }

    }
}