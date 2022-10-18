using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Order_Management_System.Services;

namespace Order_Management_System.Models
{
    public class Clientele
    {
        public string ClientID { get; set; }
        public string ClientName { get; set; }
        public string ClientSurname { get; set; }
        public string ClientAddressBlock { get; set; }
        public Clientele()
        {

        }
        public Clientele(string clientID, string clientName, string clientSurname, 
            string clientAddressBlock)
        {
            ClientID = clientID;
            ClientName = clientName;
            ClientSurname = clientSurname;
            ClientAddressBlock = clientAddressBlock;
        }
    }
}
