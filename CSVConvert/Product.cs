using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSVConvert
{
    class Product
    { // DTO object created for quick transfer of data
        public string pid { get; set; }
        public string productId { get; set; }
        public string manufacturerName { get; set; }
        public string manufacturerPN { get; set; }
        public string cost { get; set; }
        public string coo { get; set; }
        public string description { get; set; }
        public string upc { get; set; }
        public string uom { get; set; }

        public Product(string pid, string productId, string manufacturerName, string manufacturerPN, string cost, string coo, string description, string upc, string uom)
        {
            this.pid = pid;
            this.productId = productId;
            this.manufacturerName = manufacturerName;
            this.manufacturerPN = manufacturerPN;
            this.cost = cost;
            this.coo = coo;
            this.description = description;
            this.upc = upc;
            this.uom = uom;
        }

        public Product()
        {
        }

        public override string ToString()
        {
            return JsonConvert.SerializeObject(this);
        }

    }

}
