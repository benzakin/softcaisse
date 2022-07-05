using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaisseWorksheet
{
    public class Article
    {
        [Key]
        public int Article_Id { get; set; }
        public int Article_Rapport_Num { get; set; }

        public int Article_Nb { get; set; }
        public double Article_Price { get; set; }
        public string Article_Name { get; set; }
        public DateTime Article_Date { get; set; }

    }
}
