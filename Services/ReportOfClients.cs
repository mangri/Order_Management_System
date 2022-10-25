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
            LoadClientele loadclients = new LoadClientele();
            List<Clientele> clients = loadclients.RetrieveClientListFromExcelSheet();

            LoadOrders loadOrders = new LoadOrders();
            List<ReceivedOrderFull> orders = loadOrders.RetrieveOrderListFromExcelSheet();
            List<string> ListOfUnpaidOrderIDs = loadOrders.RetrieveUnpaidOrderIDsFromExcelSheet();

            string pathCustomers = "C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
                "Order_Management_System\\Reports\\CustomerSummary\\";
            string fileNameCustomers = "Report_Customers_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
            FileInfo reportCustomers = new FileInfo(pathCustomers + fileNameCustomers);
            FileStream fsCustomers = reportCustomers.Create();
            fsCustomers.Close();

            Console.WriteLine(" ****/ REPORT ON CUSTOMERS /****");
            Console.WriteLine(" ----------------------------------------------------------------");
            string customersTitle = String.Format(" {0, -15}| {1, -13}| {2, -19}| {3:0.00} ",
                "Customer ID", "First name", "Last name", "Unpaid bills");
            Console.WriteLine(customersTitle);
            Console.WriteLine(" ----------------------------------------------------------------");
            using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
            {
                sw.WriteLine(" ****/ REPORT ON CUSTOMERS /****");
                sw.WriteLine(" ----------------------------------------------------------------");
                sw.WriteLine(customersTitle);
                sw.WriteLine(" ----------------------------------------------------------------");
            }
            int unpaidBills;
            foreach (var client in clients)
            {
                unpaidBills = 0;
                for(int i = 0; i < ListOfUnpaidOrderIDs.Count; i++)
                {
                    if(orders.Where(c => c.OrderID == ListOfUnpaidOrderIDs[i]).First().ClientID == client.ClientID)
                    {
                        unpaidBills++;
                    }
                }
                
                string customerOutput = String.Format(" {0, -15}| {1, -13}| {2, -19}| {3} ",
                    client.ClientID, client.ClientFirstName, client.ClientSurname, unpaidBills);
                Console.WriteLine(customerOutput);
                using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
                {
                    sw.WriteLine(customerOutput);
                }
            }
            Console.WriteLine(" ----------------------------------------------------------------");
            Console.WriteLine(" ****/ END OF REPORT ON CUSTOMERS /****");
            using (StreamWriter sw = File.AppendText(pathCustomers + fileNameCustomers))
            {
                sw.WriteLine(" ----------------------------------------------------------------");
                sw.WriteLine(" ****/ END OF REPORT ON CUSTOMERS/****");
            }
        }
    }
}
