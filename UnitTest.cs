using System;
using Xunit;
using Order_Management_System.Repositories;
using Order_Management_System.Models;
using ClosedXML.Excel;
using ClosedXML;
using System.Linq;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace xUnitTest
{
    public class UnitTest
    {
        [Fact]
        public void RetrieveUnpaidOrder_GetsOrderList_ReturnsTheNumber()
        {
            // --Arrange
            XLWorkbook wbook = new XLWorkbook("C:\\Users\\mangri\\source\\repos\\OrderManagementSystem\\" +
                "Order_Management_System\\Databases\\Orders.xlsx");
            IXLWorksheet ws = wbook.Worksheet(1);
            LoadOrders loadOrders = new LoadOrders();
            int expectedNumber = 0;
            for(int i = 1; i <= ws.Column("F").CellsUsed().Count(); i++)
            {
                if(ws.Cell("F" + i.ToString()).GetValue<bool>() == false)
                {
                    expectedNumber++;
                }
            }

            // --Act
            int actualNumber = 
                loadOrders.RetrieveUnpaidOrderIDsFromExcelSheet().Count;

            // --Assert
            Assert.Equal(expectedNumber, actualNumber);
        }
        [Fact]
        public void StockUpdate_ReadsExcelSheet_BringsPositiveStockValues()
        {
            // --Arrange
            LoadProducts loadProducts = new LoadProducts();
            List<ProductsInStock> orders = loadProducts.RetrieveProductListFromExcelSheet();
            int expectedNegativeStocks = 0;

            // --Act
            int actualNegativeStocks = 0;
            foreach(var order in orders)
            {
                if(order.ProductsInStockAvailable < 0)
                {
                    actualNegativeStocks++;
                }
            }
                
            // --Assert
            Assert.Equal(expectedNegativeStocks, actualNegativeStocks);
        }
        [Fact]
        public void NewClientApproval_ReadsExcelSheet_GeneratesDifferentIDsForTheSameNames()
        {
            // --Arrange
            LoadClientele loadClientele = new LoadClientele();
            List<Clientele> clients = loadClientele.RetrieveClientListFromExcelSheet();
            List<string> lastNames = new List<string>();
            List<string> clientIDs = new List<string>();
            foreach (var client in clients)
            {
                lastNames.Add(client.ClientSurname);
                clientIDs.Add(client.ClientID);
            }
            bool expectedNameDublicatesButNotIdDublicates = true;

            // --Act
            
            bool actualNameDublicatesButNotIdDublicates = 
                (lastNames.Distinct().Count() < clientIDs.Distinct().Count());

            // --Assert
            Assert.Equal(expectedNameDublicatesButNotIdDublicates, actualNameDublicatesButNotIdDublicates);
        }
    }
}
