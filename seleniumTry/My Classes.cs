using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SharabiProject
{
    public class CropItem // for first row in csv
    {
        public string CropName { get; set; }
        public int CropPrice { get; set; }
        public int CropIndex { get; set; }
        public bool IsIncludeInMessage { get; set; }

    }
    public class Customer
    {// for all csv row. this mean the customer properties
        public string FirstName { get; set; }
        public string LastName { get; set; }
        public string PhoneNumber { get; set; }
        public string Time { get; set; }
        //public int Price { get; set; }
        public int Price { get; set; }
        public string[] RowData { get; set; }
        public int CsvRowIndex { get; set; }
      
    }
}
