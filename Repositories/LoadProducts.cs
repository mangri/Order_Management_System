using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using Order_Management_System.Models;


namespace Order_Management_System.Repositories
{
   public class LoadProducts
    {
        XLWorkbook wbook = new XLWorkbook("C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
            "Order_Management_System\\Databases\\Stock.xlsx");
        public List<ProductsInStock> products { get; private set; }
        public List<string> fullListOfProductIDs { get; private set; }
        public LoadProducts()
        {
            IXLWorksheet ws = wbook.Worksheet(1);
            products = new List<ProductsInStock>();
            fullListOfProductIDs = new List<string>();
            int nRow = ws.Column("A").CellsUsed().Count();
            for (int i = 1; i <= nRow; i++)
            {
                products.Add(new ProductsInStock(
                    ws.Cell("A" + i.ToString()).GetValue<string>(),
                    ws.Cell("B" + i.ToString()).GetValue<string>(),
                    ws.Cell("C" + i.ToString()).GetValue<decimal>(),
                    ws.Cell("D" + i.ToString()).GetValue<int>())
                    );
                fullListOfProductIDs.Add(ws.Cell("A" + i.ToString()).GetValue<string>());
            }
        }
        public List<ProductsInStock> RetrieveProductListFromExcelSheet()
        {
            return products;
        }
        public bool UpdateProductListInExcelSheet()
        {
            IXLWorksheet ws = wbook.Worksheet(1);
            string userInput = "To be acquired";
            Console.WriteLine("Enter 12-digits product number if available");
            var userInputTask = Task.Run(() =>
            {
                userInput = Console.ReadLine();
            });
            userInputTask.Wait(10000);
            if (userInput.All(char.IsDigit) &&
                userInput.Count() == 12 &&
                fullListOfProductIDs.Contains(userInput))
            {
                int productIndex = products.IndexOf(
                    products.Where(c => c.ProductID == userInput).First());
                // Product price
                Console.WriteLine($"Enter the new price of {products[productIndex].ProductDescription}");
                decimal updatedPrice;
                bool priceValidation = decimal.TryParse(Console.ReadLine(), out updatedPrice);
                if (priceValidation)
                {
                    ws.Cell("C" + (productIndex + 1).ToString()).SetValue(updatedPrice);
                    Console.WriteLine("The new price was accepted");
                }
                else
                {
                    Console.WriteLine("Wrong price format");
                    return false;
                }
                // Product amount
                Console.WriteLine($"Enter the new number of items in stock");
                int updatedNumber;
                bool numberValidation = Int32.TryParse(Console.ReadLine(), out updatedNumber);
                if (numberValidation)
                {
                    ws.Cell("D" + (productIndex + 1).ToString()).SetValue(updatedNumber);
                    Console.WriteLine("The new number of items was accepted");
                }
                else
                {
                    Console.WriteLine("Wrong number format");
                    return false;
                }
                wbook.Save(true);
                return true;
            }
            else
            {
                Console.WriteLine("Product number was not found. Add new product? y/n");
                string ifContinue = Console.ReadLine();
                if (ifContinue == "y" || ifContinue == "Y")
                {
                    Console.WriteLine("Please enter the name of new product");
                    string newProductName = Console.ReadLine();
                    Console.WriteLine($"Please enter the price for {newProductName}");
                    decimal newProductPrice;
                    bool newProductPriceValidation = decimal.TryParse(Console.ReadLine(), out newProductPrice);
                    Console.WriteLine($"Please enter the number of items in stock for {newProductName}");
                    int newProductNumber;
                    bool newProductNumberValidation = Int32.TryParse(Console.ReadLine(), out newProductNumber);
                    Random random = new Random();
                    string[] newProductID = new string[12];
                    for(int i = 0; i < 12; i++)
                    {
                        newProductID[i] = random.Next(0, 10).ToString();
                    }
                    if(newProductPriceValidation && newProductNumberValidation)
                    {
                        int nRow = ws.Column("A").CellsUsed().Count();
                        ws.Cell("A" + (nRow + 1).ToString()).SetValue(string.Join("", newProductID));
                        ws.Cell("B" + (nRow + 1).ToString()).SetValue(newProductName);
                        ws.Cell("C" + (nRow + 1).ToString()).SetValue(newProductPrice);
                        ws.Cell("D" + (nRow + 1).ToString()).SetValue(newProductNumber);
                        wbook.Save(true);
                        return true;
                    }
                    else
                    {
                        Console.WriteLine("The format of numbers was not accepted");
                        return false;
                    }
                }
                else
                {
                    return false;
                }
            }
        }
    }
}
