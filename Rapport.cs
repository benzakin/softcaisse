using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CaisseWorksheet
{
    public class Rapport
    {
        [Key]
        public int Rapport_Id { get; set; }
        public int Rapport_Num { get; set; }
        public DateTime Rapport_Date { get; set; }

        public double Rapport_Total { get; set; }
        public double Rapport_TVA_10 { get; set; }
        public double Rapport_TVA_20 { get; set; }
        public double Rapport_TVA_55 { get; set; }
        public double Rapport_TVA_TOTAL { get; set; }
        public double Rapport_Espece { get; set; }
        public double Rapport_CB { get; set; }
        public double Rapport_TR { get; set; }
        public double Rapport_Uber { get; set; }
        public double Rapport_Stripe { get; set; }
        public double Rapport_Cheque { get; set; }
        public int Rapport_Couvert { get; set; }

    }
}
