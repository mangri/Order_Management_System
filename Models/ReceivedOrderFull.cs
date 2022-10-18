using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Order_Management_System.Models
{
    public class ReceivedOrderFull
    {
        public string OrderID { get; set; }
        public string ClientID { get; set; }
        public string ProductNumber { get; set; }
        public int ProductAmount { get; set; }
        public decimal ProductPrice { get; set; }
        public bool PaymentReceived { get; set; }
        public DateTime OrderSubmissionTime { get; set; }
        public ReceivedOrderFull()
        {

        }
        public ReceivedOrderFull(string orderID, string clientID, string productNumber,
            int productAmount, decimal productPrice, bool paymentReceived, DateTime orderSubmissionTime)
        {
            OrderID = orderID;
            ClientID = clientID;
            ProductNumber = productNumber;
            ProductAmount = productAmount;
            ProductPrice = productPrice;
            PaymentReceived = paymentReceived;
            OrderSubmissionTime = orderSubmissionTime;
        }
    }
}
