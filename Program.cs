using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using Order_Management_System.Models;

namespace Order_Management_System
{
    internal class Program
    {
        static void Main(string[] args)
        {
            while(true)
            {
                int furtherAction = AskingUserForInput();
                switch(furtherAction)
                    {
                        case 49:

                            break;
                        case 50:

                            break;
                        case 51:

                            break;
                        case 52:

                            break;
                        case 53:
                        
                        break;
                    }
                
            }
            Console.WriteLine("")

            // Another request? y/n
            // each of them has a method
            // place excel files in a root folder for Visual Studio to seee those
            
            List<Clientele> clients = new List<Clientele>();
            string path = @"C:\Users\mangri\source\repos\OrderManagementSystem\Order_Management_System\bin\Debug\net5.0\";
            string fileNameClients = "Clientele.xlsx";
            using var wbook = new XLWorkbook(path + fileNameClients);
            var ws = wbook.Worksheet(1);
            int nRow = ws.Column("A").CellsUsed().Count();
            for (int i = 1; i <= nRow; i++)
            {
                    clients.Add(new Clientele(
                        ws.Cell("A" + i.ToString()).GetValue<string>(),
                        ws.Cell("B" + i.ToString()).GetValue<string>(),
                        ws.Cell("C" + i.ToString()).GetValue<string>(),
                        ws.Cell("D" + i.ToString()).GetValue<string>())
                        );
            }
            foreach (Clientele client in clients)
            {
                Console.WriteLine(client.ClientAddressBlock);
            }

            // Enter a new client
            // New order
            // switch for full report or particular client
        }
        private static int AskingUserForInput()
        {
            for(int i = 1; i < 6; i++)
            {
                Console.WriteLine("Choose one of the following options:");
                Console.WriteLine("[1] Add new client");
                Console.WriteLine("[2] Place new order");
                Console.WriteLine("[3] Update the stock");
                Console.WriteLine("[4] Get reports");
                Console.WriteLine("[5] Exit");
                int furtherAction = (int)Console.ReadLine();
                if(48 < furtherAction < 54)
                {
                    return furtherAction;
                    break;
                }
                else if(i < 4)
                {
                    Console.Clear();
                    Console.WriteLine($"No such option. Try again ({5 - i} attempt(s) left)");
                    continue;
                }
                else if(i == 4)
                {
                    Console.Clear();
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("The last attempt before closing the application");
                    Console.ResetColor();
                    continue;
                }
                
            }
            

        }
    }
}
