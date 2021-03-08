using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WPF_InventoryListCreator.Models
{
    class InventoryItem
    {
        public string Date { get; set; }
        public string ArticleNr { get; set; }
        public string BatchID { get; set; }
        public string Quantity { get; set; }
        public string Unit { get; set; }
        public Article article { get; set; }
    }
}
