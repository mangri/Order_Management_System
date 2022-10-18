using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Order_Management_System.Services
{
    public class ClientOrderDebts
    {
        public List<string> UnpaidOrderIDs { get; set; }
        public decimal ClientDebtTotal { get; set; }

        public ClientOrderDebts(string clientID)
        {
            UnpaidOrderIDs = RetrieveUnpaidOrders(clientID);
            ClientDebtTotal = DebtCalculator();
        }
        public List<string> RetrieveUnpaidOrders(string clientID)
        {
            return null;
        }
        public decimal DebtCalculator()
        {
            return 0.00M;
        }
    }
}
