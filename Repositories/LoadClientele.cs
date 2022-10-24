using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using Order_Management_System.Models;

namespace Order_Management_System.Repositories
{
    public class LoadClientele
    {
        XLWorkbook wbook = new XLWorkbook("C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
            "Order_Management_System\\Databases\\Clientele.xlsx");
        public List<Clientele> clients { get; private set; }
        public List<string> fullListOfClientIDs { get; private set; }
        public LoadClientele()
        {
            IXLWorksheet ws = wbook.Worksheet(1);
            clients = new List<Clientele>();
            fullListOfClientIDs = new List<string>();
            int nRow = ws.Column("A").CellsUsed().Count();
                for (int i = 1; i <= nRow; i++)
                {
                    clients.Add(new Clientele(
                        ws.Cell("A" + i.ToString()).GetValue<string>(), 
                        ws.Cell("B" + i.ToString()).GetValue<string>(), 
                        ws.Cell("C" + i.ToString()).GetValue<string>())
                        );
                    fullListOfClientIDs.Add(ws.Cell("A" + i.ToString()).GetValue<string>());
                }
        }
        public List<Clientele> RetrieveClientListFromExcelSheet()
        {
            return clients;
        }
        public bool AddNewClientToExcelSheet()
        {
            Console.WriteLine("Enter the last name:");
            string newClientSurname = Console.ReadLine();
            Console.WriteLine("Enter the first name:");
            string newClientFirstName = Console.ReadLine();
            Random rnd = new Random();
            foreach(var client in clients)
            {
                if(String.Equals(client.ClientSurname.ToLower(), newClientSurname.ToLower()) &&
                    String.Equals(client.ClientFirstName.ToLower(), newClientFirstName.ToLower()))
                {
                    Console.WriteLine($"At least one client has first name " +
                        $"'{newClientFirstName}' and surname '{newClientSurname}'");
                    Console.WriteLine("Proceed with user initiation anyway? y/n");
                    string ifContinue = Console.ReadLine();
                    if (ifContinue != "y" && ifContinue != "Y")
                    {
                        return false;
                    }
                    else
                    {
                        break;
                    }
                }
            }
            for( ; ; )
            {
                string newClientID = newClientFirstName.Substring(0, 3).ToLower() + 
                    rnd.Next(1000, 10000).ToString() + 
                    newClientSurname.Substring(0, 3).ToLower();
                if (!fullListOfClientIDs.Contains(newClientID))
                {
                    IXLWorksheet ws = wbook.Worksheet(1);
                    ws.Cell("A" + (ws.Column("A").CellsUsed().Count() + 1).ToString()).SetValue(
                        newClientID);
                    ws.Cell("B" + (ws.Column("B").CellsUsed().Count() + 1).ToString()).SetValue(
                        char.ToUpper(newClientFirstName[0]) + newClientFirstName.Substring(1).ToLower());
                    ws.Cell("C" + (ws.Column("C").CellsUsed().Count() + 1).ToString()).SetValue(
                        char.ToUpper(newClientSurname[0]) + newClientSurname.Substring(1).ToLower());
                    wbook.Save(true);
                    return true;
                }
            }
        }
    }
}
