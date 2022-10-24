using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Order_Management_System.Models;
using Order_Management_System.Repositories;

namespace Order_Management_System.Services
{
    public class ReportOfClients
    {
        public void GenerateReportForCustomers()
        {
            LoadClientele loadOrders = new LoadClientele();
            List<Clientele> clients = loadOrders.RetrieveClientListFromExcelSheet();

            string pathCustomers = "C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
                "Order_Management_System\\Reports\\CustomerSummary\\";
            string fileNameCustomers = "Report_Customers_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
            FileInfo reportCustomers = new FileInfo(pathCustomers + fileNameCustomers);
            FileStream fsCustomers = reportCustomers.Create();
            fsCustomers.Close();

            Console.WriteLine(" ****/ REPORT ON CUSTOMERS /****");
            Console.WriteLine(" --------------------------------------------------------");
            string customersTitle = String.Format(" {0, -15}| {1, -15}| {2, -20}| {3:0.00} ",
                "Customer ID", "Client first name", "Client last name", "Client debt");
            Console.WriteLine(customersTitle);
            Console.WriteLine(" --------------------------------------------------------");
            using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
            {
                sw.WriteLine(" ****/ REPORT ON CUSTOMERS /****");
                sw.WriteLine(" --------------------------------------------------------");
                sw.WriteLine(customersTitle);
                sw.WriteLine(" --------------------------------------------------------");
            }
            decimal clientDebt;
            foreach (var client in clients)
            {
                clientDebt = 0;
                string customerOutput = String.Format(" {0, -15}| {1, -15}| {2, -20}| {3:0.00} ",
                    client.ClientID, client.ClientFirstName, client.ClientSurname, clientDebt);
                Console.WriteLine(customerOutput);
                using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
                {
                    sw.WriteLine(customerOutput);
                }
            }
            Console.WriteLine(" --------------------------------------------------------");
            Console.WriteLine(" ****/ END OF REPORT ON CUSTOMERS /****");
            using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
            {
                sw.WriteLine(" --------------------------------------------------------");
                sw.WriteLine(" ****/ END OF REPORT ON CUSTOMERS/****");
            }
        }
    }
}
