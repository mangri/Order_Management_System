using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Order_Management_System.Models
{
    public class ProductsInStock
    {
        public string ProductID { get; set; }
        public string ProductDescription { get; set; }
        public decimal ProductPrice { get; set; }
        public int ProductsInStockAvailable { get; set; }
        public ProductsInStock()
        {

        }
        public ProductsInStock(string productID, string productDescription,
            decimal productPrice, int productsInStockAvailable)
        {
            ProductID = productID;
            ProductDescription = productDescription;
            ProductPrice = productPrice;
            ProductsInStockAvailable = productsInStockAvailable;
        }
    }
}
