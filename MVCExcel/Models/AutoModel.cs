using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MVCExcel.Models
{
    public class AutoModel
    {
        public string Naam { get; set; }
        public int AantalWielen { get; set; }

        public string Message { get; set; }

        public AutoModel()
        {
            Message = "Started";
        }
    }
}