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
                AskingUserForInput();
            }
            Console.WriteLine("Chose one of the following options:");
            Console.WriteLine("[1] Add new client");
            Console.WriteLine("[2] Place new order");
            Console.WriteLine("[3] Update the stock");
            Console.WriteLine("[4] Get reports");
            Console.WriteLine("[5] Exit");
            // Chose one of the following options
            // switch for: Add new client, Place new order, Update the stock, Get reports (new file for each report), Exit.
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
    }
}
