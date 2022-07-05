using CaisseWorksheet;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace GmailQuickstart
{

    class Program
    {
        static string br = "<br>";

        static void Main(string[] args)
        {

            //generate Exls
            GenerateExlsByMonth();

            IEnumerable<string> emailList;
            MailRepository mailRepository = null;
            try
            {
                mailRepository = new MailRepository(
                           "imap.online.net",
                           143,
                           false,
                           "eric@la-marmite-enchantee.fr",
                           "@Marmite06"
                       );

                //recupere les mails
                emailList = mailRepository.GetAllMails();
            }
            catch (Exception exc)
            {
                emailList = getTestEmail();
            }

            foreach (var email in emailList)
            {
                if (!string.IsNullOrEmpty(email))
                {
                    //traite les mails 
                    Console.WriteLine(email);
                    InsertEmailInDB(email);
                }
            }

            //archive les mails 
            mailRepository.GetArchiveMailsAsync("archive");

        }

        private static void GenerateExlsByMonth()
        {
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            // Creating an instance
            // of ExcelPackage
            ExcelPackage excel = new ExcelPackage();

            
            List<Rapport> rapports = new List<Rapport>();
            using (var context = new CaisseContext())
            {
                rapports = context.Rapport.OrderBy(x => x.Rapport_Num).ToList();
            }

            string day = rapports.First().Rapport_Date.ToString("MMMM 2022");
            // name of the sheet
            var workSheet = excel.Workbook.Worksheets.Add(day);

            // setting the properties
            // of the work sheet 
            workSheet.TabColor = System.Drawing.Color.Black;
            workSheet.DefaultRowHeight = 12;

            // Setting the properties
            // of the first row
            workSheet.Row(1).Height = 20;
            workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            workSheet.Row(1).Style.Font.Bold = true;

            // Header of the Excel sheet
            int i = 1;
            // workSheet.Cells[1, i++].Value = "S.No";
            // workSheet.Cells[1, i++].Value = "Id";
            workSheet.Cells[1, i++].Value = "Date";
            workSheet.Cells[1, i++].Value = "Jour";
            workSheet.Cells[1, i++].Value = "Num";
            workSheet.Cells[1, i++].Value = "Total";
            workSheet.Cells[1, i++].Value = "CB";
            workSheet.Cells[1, i++].Value = "TVA 10";
            workSheet.Cells[1, i++].Value = "TVA 20";
            workSheet.Cells[1, i++].Value = "TVA 5.5";
            workSheet.Cells[1, i++].Value = "Espece";
            workSheet.Cells[1, i++].Value = "TR";
            workSheet.Cells[1, i++].Value = "Uber";
            workSheet.Cells[1, i++].Value = "Stripe";
            workSheet.Cells[1, i++].Value = "Cheque";
            workSheet.Cells[1, i++].Value = "Couvert";

            // Inserting the article data into excel
            // sheet by using the for each loop
            // As we have values to the first row 
            // we will start with second row
            int recordIndex = 2;



            foreach (var rapport in rapports)
            {
                i = 1;
                //workSheet.Cells[recordIndex, 1].Value = (recordIndex - 1).ToString();
                // workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Id;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Date.ToString("dd");
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Date.ToString("ddd");
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Num;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Total;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_CB;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_TVA_10;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_TVA_20;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_TVA_55;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Espece;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_TR;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Uber;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Stripe;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Cheque;
                workSheet.Cells[recordIndex, i++].Value = rapport.Rapport_Couvert;


                recordIndex++;
            }

            i = 4;
            int line = recordIndex - 1;
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(D2:D" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(E2:E" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(F2:F" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(G2:G" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(H2:H" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(I2:I" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(J2:J" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(K2:K" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(L2:L" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(M2:M" + line + ")";
            workSheet.Cells[recordIndex, i++].Formula = "=SUM(N2:N" + line + ")";

            workSheet.Row(recordIndex).Style.Font.Bold = true;


            for (int nb = 1; nb < i+4; nb++)
                workSheet.Column(nb).AutoFit();

            workSheet.PrinterSettings.Orientation = eOrientation.Landscape;

            // file name with .xlsx extension 
            string p_strPath = "c:\\pc\\perso\\marmite\\" + DateTime.Now.ToString("MMMM") + "_Caisse.xlsx";

            if (File.Exists(p_strPath))
                File.Delete(p_strPath);

            // Create excel file on physical disk 
            FileStream objFileStrm = File.Create(p_strPath);
            objFileStrm.Close();

            // Write content to excel file 
            File.WriteAllBytes(p_strPath, excel.GetAsByteArray());
            //Close Excel package
            excel.Dispose();
        }

        public static void InsertEmailInDB(string email)
        {

            using (var context = new CaisseContext())
            {
                int rapportNum = 0;
                DateTime rapportDate = new DateTime();
                double rapportTotal = 0;
                double rapportTva10 = 0;
                double rapportTva20 = 0;
                double rapportTva55 = 0;
                decimal rapportTvaTotal = 0;
                double rapportEspece = 0;
                double rapportCB = 0;
                double rapportTR = 0;
                double rapportUber = 0;
                double rapportStripe = 0;
                double rapportCheque = 0;
                int rapportCouvert = 0;

                string[] lines = email.Split("\r\n");

                var listlines = new List<string>(lines);

                rapportNum = GetRapportNum(lines);
                rapportDate = GetRapportDate(lines);

                GetArticle(listlines, rapportNum, rapportDate);

                rapportTotal = GetRapportTotal(listlines, "Total Paiements");
                rapportTva10 = GetRapportTva10(listlines, "10,00 %");
                rapportTva20 = GetRapportTva20(listlines, "20,00 %");
                rapportTva55 = GetRapportva55(listlines, "5,50 %");
                rapportTvaTotal = (decimal)rapportTva10 + (decimal)rapportTva20 + (decimal)rapportTva55;
                rapportEspece = GetRapportEspece(listlines, "Total Especes Perues");
                rapportCB = SearchCB(listlines, "Carte Bleue"); 
                rapportTR = SearchCB(listlines, "Ticket Restauran");// GetRapportTR(listlines, "Ticket Restauran");
                rapportUber = SearchCB(listlines, "UberEat");//  GetRapportUber(listlines, "UberEat");
                rapportStripe = SearchCB(listlines, "Stripe");//  GetRapportStripe(listlines, "Stripe");
                rapportCouvert = GetRapportCouvert(listlines, "Couverts");
                rapportCheque = SearchCB(listlines, "CHEQUE");



                var std = new Rapport()
                {
                    Rapport_Num = rapportNum,
                    Rapport_Date = rapportDate,
                    Rapport_Total = rapportTotal,
                    Rapport_TVA_10 = rapportTva10,
                    Rapport_TVA_20 = rapportTva20,
                    Rapport_TVA_55 = rapportTva55,
                    Rapport_TVA_TOTAL = double.Parse(rapportTvaTotal.ToString()),
                    Rapport_Espece = rapportEspece,
                    Rapport_CB = rapportCB,
                    Rapport_TR = rapportTR,
                    Rapport_Uber = rapportUber,
                    Rapport_Stripe = rapportStripe,
                    Rapport_Cheque = rapportCheque,
                    Rapport_Couvert = rapportCouvert

                };

                context.Rapport.Add(std);
                context.SaveChanges();
            }

        }

        private static int GetRapportCouvert(List<string> lines, string searchVal)
        {
            return int.Parse(SearchCouvert(lines, searchVal).ToString());
        }

        //private static double GetRapportStripe(List<string> lines, string searchVal)
        //{
        //    return SearchCB(lines, searchVal);
        //}

        //private static double GetRapportUber(List<string> lines, string searchVal)
        //{
        //    return SearchCB(lines, searchVal);
        //}

        //private static double GetRapportTR(List<string> lines, string searchVal)
        //{
        //    return Search(lines, searchVal);
        //}

        //private static double GetRapportCB(List<string> lines, string searchVal)
        //{
        //    return SearchCB(lines, searchVal);
        //}

        private static double GetRapportEspece(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }

        private static double GetRapportva55(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }
        private static double GetRapportTvaTotal(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }


        private static double GetRapportTva20(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }

        private static double GetRapportTva10(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }

        private static double Search(List<string> lines, string searchVal)
        {
            double total = 0;

            string currentLine = lines.FirstOrDefault<string>(x => x.Contains(searchVal));
            if (!string.IsNullOrEmpty(currentLine))
            {
                int substr1 = currentLine.IndexOf(searchVal) + searchVal.Length;
                int substr2 = currentLine.LastIndexOf(br) - substr1;

                string strTotal = currentLine.Substring(substr1, substr2);
                total = double.Parse(strTotal.Trim());
            }
            return total;
        }
        private static double SearchCouvert(List<string> lines, string searchVal)
        {
            double total = 0;

            string currentLine = lines.FirstOrDefault<string>(x => x.Contains(searchVal));
            if (!string.IsNullOrEmpty(currentLine))
            {
                int substr1 = searchVal.Length;

                string strTotal = currentLine.Substring(substr1).Trim().Substring(0, 3);
                total = double.Parse(strTotal.Trim());
            }
            return total;
        }
        private static double SearchCB(List<string> lines, string searchVal)
        {
            double total = 0;

            string currentLine = lines.FirstOrDefault<string>(x => x.Contains(searchVal));
            if (!string.IsNullOrEmpty(currentLine))
            {
                int substr1 = currentLine.Length - 10;
                //int substr2 = currentLine.LastIndexOf(br) - substr1;

                string strTotal = currentLine.Substring(substr1).Replace("<br>", "");
                total = double.Parse(strTotal.Trim());
            }
            return total;
        }

        private static double GetRapportTotal(List<string> lines, string searchVal)
        {
            return Search(lines, searchVal);
        }

        private static int GetRapportNum(string[] lines)
        {
            string zrapport = "ZRAPPORT# ";
            try
            {
                int substr1 = lines[0].IndexOf(zrapport) + zrapport.Length;
                int substr2 = lines[0].LastIndexOf(br) - substr1;
                return int.Parse(lines[0].Substring(substr1, substr2));
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        private static DateTime GetRapportDate(string[] lines)
        {

            try
            {
                return DateTime.Parse(lines[2].Substring(0, 8) + " " + lines[2].Substring(31, 8));
            }
            catch (Exception exc)
            {
                throw exc;
            }
        }

        private static void GetArticle(List<string> listlines, int articleRapportNum, DateTime articleDate)
        {
            List<string> listCategory = new List<string>() { "A la carte", "Entree unite", "Menu", "Desserts", "Boissons", "Sans alcool", "Jour", "Woks","Bobun" };
            //get Article
            for (int nb = 0; nb < 5; nb++)
            {
                listlines.RemoveAt(0);
            }

            int nbArticle = 0;
            string articleName = "";
            double price = 0;
            foreach (var line in listlines)
            {
                if (line.IndexOf("x") == -1)
                    break;
                nbArticle = int.Parse(line.Substring(0, line.IndexOf("x")).Trim());
                int nbLength = line.Length;
                string strPrice = "";
                strPrice = line.Substring(nbLength - 10).Trim();
                articleName = line.Substring(line.IndexOf("x") + 1).Trim();
                articleName = articleName.Replace(strPrice, "").Trim();
                strPrice = strPrice.Replace("<br>", "");

                price = double.Parse(strPrice);

                if (!listCategory.Contains(articleName))
                {
                    using (var context = new CaisseContext())
                    {
                        var art = new Article()
                        {
                            Article_Rapport_Num = articleRapportNum,
                            Article_Nb = nbArticle,
                            Article_Price = price,
                            Article_Name = articleName,
                            Article_Date = articleDate

                        };

                        context.Article.Add(art);
                        context.SaveChanges();
                    }
                }
            }
        }


        private static IEnumerable<string> getTestEmail()
        {
            string result = @"
<div dir=ltr><br><br><div class=gmail_quote><div dir=ltr class=gmail_attr>---------- Forwarded message ---------<br>De : <strong class=gmail_sendername dir=auto>La Marmite Enchantee</strong> <span dir=auto>&lt;<a href=mailto: rapportnova @jdc.fr>rapportnova@jdc.fr</a>&gt;</span><br>Date: mar. 28 juin 2022 à 22:59<br>Subject: Rapport Email de caisse<br>To:  &lt;<a href=mailto: marmiteenchantee @gmail.com>marmiteenchantee@gmail.com</a>&gt;<br></div><br><br>ZRAPPORT# 895<br>
    < br >
    25 / 06 / 22    POS #01        13:25:12<br>
________________Z RAPPORT________________<br>

       ARTICLES.JOUR < br >
   7 x A la carte        28,50 < br >
     3 x A la carte      22,50 < br >
      2 x Assiette decouv   22,00 < br >
      1 x Suppl.                0,50 < br >
     4 x Entree unite      6,00 < br >
      1 x Samossa chevre    1,50 < br >
      1 x Samossa boeuf    1,50 < br >
      2 x Brochette yakitor  3,00 < br >
   3 x Menu           61,50 < br >
      1 x Formule entree,bo  17,50 < br >
      2 x REPAS COMPLET    44,00 < br >
   3 x Desserts         15,00 < br >
     3 x Desserts       15,00 < br >
      1 x Creme brulee     5,00 < br >
      1 x Tarte du jour    5,00 < br >
      1 x Dessert jr      5,00 < br >
   6 x Boissons         17,00 < br >
     1 x Boissons alcool    4,00 < br >
      1 x BIERE LOCALE     4,00 < br >
     5 x Sans alcool      13,00 < br >
      1 x Eau 1L        3,00 < br >
      3 x Coca 0        7,50 < br >
      1 x JUS         2,50 < br >
   2 x Jour           27,00 < br >
      2 x Plat du jour    27,00 < br >
   1 x Woks           15,00 < br >
      1 x Wok boeuf + boul   15,00 < br >
     5 x BOBUN           55,00 < br >
        5 x Bo Bun       55,00 < br >
  __________________________________________<br>
  < br >
                 TOTAL 219,00 < br >
  < br >
  ZRAPPORT# 895<br>
< br >
  25 / 06 / 22    POS #01        13:25:12<br>
________________Z RAPPORT________________<br>

       GROUPES.JOUR<br>

         VIDE<br>
< br >
ZRAPPORT# 895<br>
< br >
25 / 06 / 22    POS #01        13:25:12<br>
________________Z RAPPORT________________<br>

        TVA.JOUR<br>
(1)  10,00 % 18,36 < br >
(2)  20,00 % 0,67 < br >
(3)  5,50 % 0,68 < br >
__________________________________________<br>
< br >
TOTAL 19,71 < br >
< br >
ZRAPPORT# 895<br>
< br >
25 / 06 / 22    POS #01        13:25:12<br>
________________Z RAPPORT________________<br>
      TVA OFFERTS.JOUR<br>

         VIDE<br>
< br >
LA MARMITE ENCHANTEE<br>
20 ROUTE DE GOURDON < br >
06740 CHATEAUNEUF DE GRASSE FRANCE<br>
83170001800011 5610C < br >
09831700018 < br >
------------------------------------------< br >
        FINANCIER < br >
------------------------------------------< br >
ZRAPPORT#      000895<br>
< br >
JOUR<br>
< br >
25 / 06 / 22    POS #01        13:25:12<br>
002 TICKET<br>
DE :   1096 / 01  A: 1101 / 01  219,00 < br >
             _____________________ < br >
                 TOTAL 219,00 < br >
   Pay                219,00 < br >
   En Compte              0,00 < br >
   ------------------------------------------< br >
   ******************TVA * ******************< br >
            HT    TVA TTC<br>

 10,00 % 183,64   18,36  202,00 < br >
20,00 % 3,33   0,67   4,00 < br >
5,50 % 12,32   0,68   13,00 < br >
---------------------------< br >
  199,29   19,71  219,00 < br >
***************PAIEMENTS * ***************< br >
Especes               28,00 < br >
Total Especes Perues        28,00 < br >
Total Mouvements Tiroir       0,00 < br >
     ---------------------< br >
Total Tiroir            28,00 < br >
Total Especes            28,00 < br >
< br >
Carte Bleue       7    191,00 < br >
Total autres paiements       191,00 < br >
     ---------------------< br >
   Total Paiements  219,00 < br >
------------------------------------------< br >
****************COUVERTS * ***************< br >
Couverts        10     0,00 < br >
------------------------------------------< br >
__________________________________________ < br >
*****************Offert * ****************< br >
***************ANNULATION * **************< br >
ANNUL APRES ENVOI          33,50 < br >
ANNUL.APRES NOTE          66,00 < br >
ANNUL.TOTAL(2xC)         66,00 < br >
CORRECTION             31,00 < br >
------------------------------------------< br >
< br >
< br >
</ div >< br clear = all >< div >< br ></ div > -- < br >< div dir = ltr class=gmail_signature data-smartmail=gmail_signature><div dir = ltr >< div >< div dir=ltr><font size = 1 > Cécile BENZAKIN, <br>Présidente du Restaurant  :<br></font><div><font size = 1 > LA MARMITE ENCHANTEE</font></div><div><font size = 1 > 20 route de Gourdon</font></div><div><font size = 1 > 06740 CHATEAUNEUF DE GRASSE</font></div><div><i><b>04 93 09 08 22</b></i></div></div></div></div></div></div>
 
                                                                                 ";



                                    List<string> l = new List<string>();
            l.Add(result);
            return l.AsEnumerable();
        }
    }
}