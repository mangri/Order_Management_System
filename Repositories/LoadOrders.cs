using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Order_Management_System.Models;

namespace Order_Management_System.Repositories
{
    public class LoadOrders
    {
        XLWorkbook wbook = new XLWorkbook("C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
            "Order_Management_System\\Databases\\Orders.xlsx");
        public List<ReceivedOrderFull> orders { get; set; }
        public List<string> ListOfUnpaidOrderIDs { get; private set; }
        public List<string> fullListOfOrderIDs { get; private set; }
        public LoadOrders()
        {
            IXLWorksheet ws = wbook.Worksheet(1);
            orders = new List<ReceivedOrderFull>();
            ListOfUnpaidOrderIDs = new List<string>();
            fullListOfOrderIDs = new List<string>();
            int nRow = ws.Column("A").CellsUsed().Count();
            for (int i = 1; i <= nRow; i++)
            {
                orders.Add(new ReceivedOrderFull(
                    ws.Cell("A" + i.ToString()).GetValue<string>(),
                    ws.Cell("B" + i.ToString()).GetValue<string>(),
                    ws.Cell("C" + i.ToString()).GetValue<string>(),
                    ws.Cell("D" + i.ToString()).GetValue<int>(),
                    ws.Cell("E" + i.ToString()).GetValue<decimal>(),
                    ws.Cell("F" + i.ToString()).GetValue<bool>(),
                    ws.Cell("G" + i.ToString()).GetValue<DateTime>())
                    );
                fullListOfOrderIDs.Add(ws.Cell("A" + i.ToString()).GetValue<string>());
                if (orders[i - 1].PaymentReceived == false)
                {
                    ListOfUnpaidOrderIDs.Add(ws.Cell("A" + i.ToString()).GetValue<string>());
                }
            }
        }
        public List<ReceivedOrderFull> RetrieveOrderListFromExcelSheet()
        {
            return orders;
        }
        public List<string> RetrieveUnpaidOrderIDsFromExcelSheet()
        {
            return ListOfUnpaidOrderIDs;
        }
        public List<string> RetrieveFullListOfOrderIDsFromExcelSheet()
        {
            return fullListOfOrderIDs;
        }
        public bool AddNewOrderToExcelSheet()
        {
            IXLWorksheet ws = wbook.Worksheet(1);
            // Load customer data
            LoadClientele loadClientele = new LoadClientele();
            List<string> clientIDs = loadClientele.fullListOfClientIDs;
            List<Clientele> allClientInfo = loadClientele.RetrieveClientListFromExcelSheet();
            // Load product data
            LoadProducts loadProducts = new LoadProducts();
            List<string> productIDs = loadProducts.fullListOfProductIDs;
            List<ProductsInStock> allProductInfo = loadProducts.RetrieveProductListFromExcelSheet();

            Console.WriteLine("Enter clientID to confirm the order for");
            string orderingClient = Console.ReadLine();
            if (clientIDs.Contains(orderingClient))
            {
                Console.WriteLine($"Enter product code to be sold to " +
                    $"{allClientInfo.Where(c => c.ClientID == orderingClient).First().ClientSurname}");
                string requestedProduct = Console.ReadLine();
                if(productIDs.Contains(requestedProduct))
                {
                    Console.WriteLine("Both IDs were confirmed");
                    Console.WriteLine($"How many " +
                    $"{allProductInfo.Where(c => c.ProductID == requestedProduct).First().ProductDescription} " +
                    $"to reserve? " +
                    $"(Available: {allProductInfo.Where(c => c.ProductID == requestedProduct).First().ProductsInStockAvailable})");
                    int requestedAmount;
                    bool finalRequestValidation = Int32.TryParse(Console.ReadLine(), out requestedAmount);
                    if (!finalRequestValidation)
                    {
                        Console.WriteLine("Invalid input format");
                        return false;
                    }
                    else if(requestedAmount <= allProductInfo.Where(c => c.ProductID == requestedProduct).First().ProductsInStockAvailable)
                    {
                        Random random = new Random();
                        string[] newOrderDigitsPartForID = new string[6];
                        for (int i = 0; i < 6; i++)
                        {
                            newOrderDigitsPartForID[i] = random.Next(0, 10).ToString();
                        }
                        string newOrderID = allClientInfo.Where(c => c.ClientID == orderingClient).First().ClientSurname.Substring(0, 3).ToUpper() + 
                            String.Join("", newOrderDigitsPartForID);
                        decimal productPrice = allProductInfo.Where(c => c.ProductID == requestedProduct).First().ProductPrice;
                        bool paymentReceived;
                        if (random.Next(0, 2) == 0) paymentReceived = false; else paymentReceived = true;
                        if (!RetrieveFullListOfOrderIDsFromExcelSheet().Contains(newOrderID))
                        {
                            int nRow = ws.Column("A").CellsUsed().Count();
                            ws.Cell("A" + (nRow + 1).ToString()).SetValue(newOrderID);
                            ws.Cell("B" + (nRow + 1).ToString()).SetValue(orderingClient);
                            ws.Cell("C" + (nRow + 1).ToString()).SetValue(requestedProduct);
                            ws.Cell("D" + (nRow + 1).ToString()).SetValue(requestedAmount);
                            ws.Cell("E" + (nRow + 1).ToString()).SetValue(productPrice);
                            ws.Cell("F" + (nRow + 1).ToString()).SetValue(paymentReceived);
                            ws.Cell("G" + (nRow + 1).ToString()).SetValue(DateTime.Now);
                            wbook.Save(true);
                            XLWorkbook wbookStock = new XLWorkbook("C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
                                "Order_Management_System\\Databases\\Stock.xlsx");
                            IXLWorksheet wsStock = wbookStock.Worksheet(1);
                            int productIndexInTheList = allProductInfo.IndexOf(allProductInfo.Where(c => c.ProductID == requestedProduct).First());
                            wsStock.Cell("D" + (productIndexInTheList + 1).ToString()).SetValue(allProductInfo[productIndexInTheList].ProductsInStockAvailable - requestedAmount);
                            wbookStock.Save(true);
                            return true;
                        }
                        else
                        {
                            Console.WriteLine("Preventing the same order IDs to be placed. Abort!");
                            return false;
                        }
                    }
                    else
                    {
                        Console.WriteLine("Too many items requested. Stock up on goods!");
                        return false;
                    }
                }
                else
                {
                    Console.WriteLine("Requested product was not found");
                    return false;
                }
            }
            Console.WriteLine("Specified customer was not found");
            return false;
        }
    }
}
