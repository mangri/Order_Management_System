using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Order_Management_System.Repositories;
using Order_Management_System.Models;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.IO;

namespace Order_Management_System.Services
{
    public class ReportOfUnpaidOrders
    {

        public void GenerateReportForUnpaidOrders()
        {
            LoadOrders loadOrders = new LoadOrders();
            List<string> listOfUnpaidOrderIDs = loadOrders.RetrieveUnpaidOrderIDsFromExcelSheet();
            List<ReceivedOrderFull> orders = loadOrders.RetrieveOrderListFromExcelSheet();
            // Creating txt file
            string pathUnpaid = "C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\Order_Management_System\\Reports\\UnpaidOrders\\";
            string fileNameUnpaid = "Report_UnpaidOrders_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".txt";
            FileInfo reportUnpaid = new FileInfo(pathUnpaid + fileNameUnpaid);
            FileStream fsUnpaid = reportUnpaid.Create();
            fsUnpaid.Close();

            Console.WriteLine(" ****/ REPORT FOR UNPAID ORDERS /****");
            Console.WriteLine(" --------------------------------------------------------");
            string unpaidOrdersTitle = String.Format(" {0, -10}| {1, -11}| {2, -13}| {3, -7}| {4:0.00} ",
                "Order ID", "Client ID", "Product code", "Units", "Price/unit");
            Console.WriteLine(unpaidOrdersTitle);
            Console.WriteLine(" --------------------------------------------------------");
            using (StreamWriter sw = File.AppendText(pathUnpaid + fileNameUnpaid))
            {
                sw.WriteLine(" ****/ REPORT FOR UNPAID ORDERS /****");
                sw.WriteLine(" --------------------------------------------------------");
                sw.WriteLine(unpaidOrdersTitle);
                sw.WriteLine(" --------------------------------------------------------");
            }
            foreach (string unpaidOrderID in listOfUnpaidOrderIDs)
            {
                string unpaidOrdersOutput = String.Format(" {0, -10}| {1, -11}| {2, -13}|  {3, -6}| {4:0.00} ", 
                    unpaidOrderID, 
                    orders.Where(c => c.OrderID == unpaidOrderID).First().ClientID, 
                    orders.Where(c => c.OrderID == unpaidOrderID).First().ProductNumber, 
                    orders.Where(c => c.OrderID == unpaidOrderID).First().ProductAmount, 
                    orders.Where(c => c.OrderID == unpaidOrderID).First().ProductPrice);
                Console.WriteLine(unpaidOrdersOutput);
                using (StreamWriter sw = File.AppendText(pathUnpaid + fileNameUnpaid))
                {
                    sw.WriteLine(unpaidOrdersOutput);
                }
            }
            Console.WriteLine(" --------------------------------------------------------");
            Console.WriteLine(" ****/ END OF FIRST REPORT /****");
            using (StreamWriter sw = File.AppendText(pathUnpaid + fileNameUnpaid))
            {
                sw.WriteLine(" --------------------------------------------------------");
                sw.WriteLine(" ****/ END OF REPORT FOR UNPAID ORDERS/****");
            }
        }
    }
}
