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
