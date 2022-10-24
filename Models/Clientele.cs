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
        public string ClientFirstName { get; set; }
        public string ClientSurname { get; set; }
        public Clientele()
        {

        }
        public Clientele(string clientID, string clientFirstName, string clientSurname)
        {
            ClientID = clientID;
            ClientFirstName = clientFirstName;
            ClientSurname = clientSurname;
        }
    }
}
