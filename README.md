# csharp-SalesOrderEOD
This reads a specified users emails looking for an email in a date range with a specified sender and takes the most recently received email and downloads the attached .csv file. Then extracts the order confirmations from the .csv file and updates the order status in our system of matching orders to compelete while adding their consignments. An email report is sent of the orders that have been updated or unable to be updated for the support team to follow up.
