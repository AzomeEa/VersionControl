using mintaZH.Program;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace mintaZH
{
    public partial class Form1 : Form
    {
        List<mintaZH.Program.OlympicResult> results = new List<mintaZH.Program.OlympicResult>();
        Excel.Application xlApp;
        Excel.Workbook xlWB;
        Excel.Worksheet xlSheet;

        public Form1()
        {
            InitializeComponent();
            LoadData("Summer_olympic_Medals");
        }

        void LoadData(string fileName) //bemenő paramétere egy random fájl neve
        {
            using (StreamReader sr = new StreamReader(fileName, Encoding.Default))
            {
                sr.ReadLine(); // Az első sort (fejléc) eldobjuk
                while (!sr.EndOfStream)
                {
                    var line = sr.ReadLine().Split(','); //vesszők szerint tördel
                    var or = new mintaZH.Program.OlympicResult() //itt olvassa a listába
                    {
                        Year = int.Parse(line[0]),
                        Country = line[3],
                        Medals = new int[] //tömb
                        {
                            int.Parse(line[5]),
                            int.Parse(line[6]),
                            int.Parse(line[7])
                        }
                    };
                    results.Add(or);
                }
            }
        }

        void CreateYearFilter()
        {
            //LINQ lekérdezés
            var years = (from r in results //a results listából keres
                         orderby r.Year //év szerint rendez
                         select r.Year).Distinct(); //kiszűri a duplikálást
            cboxYears.DataSource = years.ToList(); //a ComboBox adatforrása a lekérdezés eredméyne
        }

        int CalculatePosition(mintaZH.Program.OlympicResult or) //a bemenete EGY OlympicResult
        {
            var betterCountryCount = 0; //számláló kezdőértékkel
            //Szűrés LINQ lekérdezéssel
            var filteredResults = from r in results //szűrjük a listás úgy...
                                  where r.Year == or.Year && r.Country != or.Country //hogy a bemenettel (tehát or) megegyezzen az év, de az ország ne
                                  select r; //válasszuk ki az adott sort

            foreach (var r in filteredResults) //végigmegyünk a szűrt lista elemein
            {
                if (r.Medals[0] > or.Medals[0]) //Az adott sor aranyainak száma nagyobb, mint a bemenet aranyainak száma.
                {
                    betterCountryCount++;
                }
                else if (r.Medals[0] == or.Medals[0])
                {
                    if (r.Medals[1] > or.Medals[1]) //Az adott sor aranyainak száma ugyanannyi, mint a bemenet aranyainak száma, de az ezüstök száma nagyobb.
                    {
                        betterCountryCount++;
                    }
                    else if (r.Medals[1] == or.Medals[1]) //Az adott sor aranyainak és ezüstjeinek száma rendre ugyanannyi, mint a bemenet aranyainak és ezüstjeinek száma, de a bronzok száma nagyobb.
                    {
                        if (r.Medals[2] > or.Medals[2])
                        {
                            betterCountryCount++;
                        }
                    }
                }
            }
            return 0; //kimeneti értéke egy szám
            return betterCountryCount + 1; //a függvény visszatér a számlálónál 1-gyel nagyobb értékkel
        }

        void CalculateOrder() //a Form1 konstruktorában, paraméter nélküli fv
        {
            foreach (var r in results)
            {
                r.Position = CalculatePosition(r); //feltölti a Position tul. értékét 
            }
        }

        private void btnExport_Click(object sender, EventArgs e) //kattintásra excelt kreál
        {
            try
            {
                xlApp = new Excel.Application();
                xlWB = xlApp.Workbooks.Add(Missing.Value);
                xlSheet = xlWB.ActiveSheet;

                // Excel feltöltése
                // ...

                xlApp.Visible = true;
                xlApp.UserControl = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

                xlWB.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
                xlWB = null;
                xlApp = null;
            }
        }

        void WriteExcel() //adatok excelbe írása
        {
            var headers = new string[] //kimeneti állomány fejléce
            {
                "Helyezés",
                "Ország",
                "Arany",
                "Ezüst",
                "Bronz"
            };
            for (int i = 0; i < headers.Length; i++)
                xlSheet.Cells[1, i + 1] = headers[i];

            var filteredResults = from r in results
                                  where r.Year == (int)cboxYears.SelectedItem
                                  orderby r.Position
                                  select r;

            var sor = 2; //2. sortól tölti fel értékkel
            foreach (var r in filteredResults)
            {
                xlSheet.Cells[sor, 1] = r.Position;
                xlSheet.Cells[sor, 2] = r.Country;

                for (int i = 0; i < 2; i++)
                {
                    xlSheet.Cells[sor, i + 3] = r.Medals[i];
                    sor++;
                }
            }
        }
    
    }
}
