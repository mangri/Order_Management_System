using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Wordprocessing;
using Order_Management_System.Models;
using Order_Management_System.Repositories;
using Order_Management_System.Services;

namespace Order_Management_System
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                char validatedAction = AskingUserForInput();
                switch (validatedAction)
                    {
                        case '1':
                        NewClientInitializer();
                            break;
                        case '2':
                        NewOrderInitializer();
                        break;
                        case '3':
                        ProductListUpdateInitializer();
                            break;
                        case '4':
                        ReportInitializer();
                            break;
                        case '5':
                        Environment.Exit(0);
                        break;

                        default: break;
                    }
                if(validatedAction == 'Z')
                {
                    Console.WriteLine("Too many attempts to connect.");
                    Console.WriteLine("The systems was not loaded."); 
                    Console.WriteLine("Try again later. Bye.");
                    break;
                }
                Console.WriteLine("Proceed with one more request? y/n");
                string ifContinue = Console.ReadLine();
                if(ifContinue == "y" || ifContinue == "Y")
                {
                    continue;
                }
                else
                {
                    break;
                }
            }
            Console.WriteLine("Exiting application..");
            Console.ReadLine();
        }
        private static char AskingUserForInput()
        {
            for(int i = 1; i < 6; i++)
            {
                Console.WriteLine("Choose one of the following options:");
                Console.WriteLine("[1] Add new client");
                Console.WriteLine("[2] Place new order");
                Console.WriteLine("[3] Update the stock");
                Console.WriteLine("[4] Get reports");
                Console.WriteLine("[5] Exit");
                string furtherAction = Console.ReadLine();
                char validatedAction;
                bool goodToGo = Char.TryParse(furtherAction, out validatedAction);
                if(goodToGo && '0' < validatedAction && validatedAction < '6')
                {
                    return validatedAction;
                }
                else if(i < 4)
                {
                    //Console.Clear();
                    Console.WriteLine($"No such option. Try again ({5 - i} attempt(s) left)");
                }
                else if(i == 4)
                {
                    //Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("The last attempt before closing the application");
                    Console.ResetColor();
                }
            }
            return 'Z';
        }
        private static void NewClientInitializer()
        {
            LoadClientele loadClientele = new LoadClientele();
            bool ifCreated = loadClientele.AddNewClientToExcelSheet();
            if(ifCreated)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Client database was successfully updated");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("New client initiation was declined");
                Console.ResetColor();
            }
        }
        private static void ProductListUpdateInitializer()
        {
            LoadProducts loadProducts = new LoadProducts();
            bool ifUpdated = loadProducts.UpdateProductListInExcelSheet();
            if (ifUpdated)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Product database was successfully updated");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Update of products in stock was declined");
                Console.ResetColor();
            }
        }
        private static void NewOrderInitializer()
        {
            LoadOrders loadOrders = new LoadOrders();
            bool ifOrderMade = loadOrders.AddNewOrderToExcelSheet();
            if (ifOrderMade)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("The order went through successfully");
                Console.ResetColor();
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("The order was not made");
                Console.ResetColor();
            }
        }
        private static void ReportInitializer()
        {
            ReportOfUnpaidOrders reportOfUnpaidOrders = new ReportOfUnpaidOrders();
            reportOfUnpaidOrders.GenerateReportForUnpaidOrders();
            Console.WriteLine();
            ReportOfClients reportOfClients = new ReportOfClients();
            reportOfClients.GenerateReportForCustomers();
            Console.WriteLine();
        }
    }
}
