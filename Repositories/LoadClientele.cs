using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using Order_Management_System.Models;

namespace Order_Management_System.Repositories
{
    public class LoadClientele
    {
        XLWorkbook wbook = new XLWorkbook("Clientele.xlsx");
        public List<Clientele> clients { get; private set; }
        public LoadClientele()
        {
            var ws = wbook.Worksheet(1);
            clients = new List<Clientele>();
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
        }
        public List<Clientele> RetrieveClientList()
        {
            return clients;
        }
        public void AddNewCliwnt()
        {
            
        }
    }
}
