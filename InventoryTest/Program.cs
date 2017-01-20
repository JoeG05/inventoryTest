using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace InventoryTest
{
    class Program
    {
        struct SalesData
        {
            
            public double Sold;
            public double OH;
            public string Code;
        };

        static void getInventory(Dictionary<string, SalesData> d)
        {
            char[] delimiterChars = { ',' };
            string[] productCodes = { "584-10", "584-30", "555-00", "580-00", "581-00", "587-00" };
            string line;
            
            // Scotch plains file location
            // StreamReader file = new StreamReader(@"C:\Dropbox\Work\Inventory\Report.txt");

            // Chatham file location
            StreamReader file = new StreamReader(@"C:\Dropbox\Work\Chatham\Report.txt");

            while ((line = file.ReadLine()) != null)
            {


                string[] words = line.Split(delimiterChars);
                if (words[0] != productCodes[0] &&
                    words[0] != productCodes[1] &&
                    words[0] != productCodes[2] &&
                    words[0] != productCodes[3] &&
                    words[0] != productCodes[4] &&
                    words[0] != productCodes[5] ||
                    words[2] == "Total $:")
                {
                    continue;
                }
                SalesData s = new SalesData();

                s.Sold = double.Parse(words[9]);
                s.OH = double.Parse(words[13]);
                s.Code = words[0];
                d.Add(words[2], s);
            }
            file.Close();
        }

        static void nameConversion(Dictionary<string, SalesData> d, string file)
        {
            Excel.Application xl = new Excel.Application();
            Excel.Workbook xlWorkbook = xl.Workbooks.Open(file);
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[2];
            Excel.Range xlRange = xlSheet.UsedRange;

            int counter = 1;
            foreach(string entry in d.Keys)
            {
                xlSheet.Cells[counter, 1] = entry;
                counter++;
            }
            xlWorkbook.Save();
            xlWorkbook.Close();
            
        }

        static void fillPerpetual(Dictionary<string, SalesData> d, string file)
        {
            Excel.Application xl = new Excel.Application();
            Excel.Workbook xlWorkbook = xl.Workbooks.Open(file);
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlSheet.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i < rowCount; i++)
            {
                string item = (string)(xlRange.Cells[i, 1] as Excel.Range).Value2;
                if (item == null)
                    continue;
                if (d.ContainsKey(item))
                {
                    xlRange.Cells[i, 3].Value2 = d[item].OH;
                }
            }
            // Scotch Plains
            // string saveAs = @"C:\Dropbox\Work\Inventory\Perpetual-WE-" + getNextSunday();

            // Chatham
            string saveAs = @"C:\Dropbox\Work\Chatham\Perpetual-WE-" + getNextSunday();
            
            xlWorkbook.SaveAs(saveAs);
            xlWorkbook.Close(file);

            Console.WriteLine("Perpetual copy completed.");

        }

        static void orderConversion(Dictionary<string, SalesData>d, string file)
        {
            Excel.Application xl = new Excel.Application();
            Excel.Workbook xlWorkbook = xl.Workbooks.Open(file);
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlSheet.UsedRange;


            foreach (var entry in d)
            {
                Console.WriteLine(entry.Key);
                if (entry.Value.Code == "584-10" || entry.Value.Code == "584-30" || entry.Value.Code == "555-00")
                    continue;

                int row = Int32.Parse(Console.ReadLine());

                if (row == 999)
                    continue;
                if (row == -1)
                {
                    xlWorkbook.Save();
                    xlWorkbook.Close(file);
                    Console.WriteLine("Complete");
                    return;
                }
                xlRange.Cells[row, 1] = entry.Key;
                

            }
            xlWorkbook.Save();
            xlWorkbook.Close(file);
            Console.WriteLine("orderConversion Finished.");
        }

        static string getNextSunday()
        {
            DateTime today = DateTime.Today;
            int daysUntilSunday = ((int)DayOfWeek.Sunday - (int)today.DayOfWeek + 7) % 7;
            DateTime nextSunday = today.AddDays(daysUntilSunday);
            return nextSunday.ToString("M");

        }

        static string getToday()
        {
            DateTime today = DateTime.Today;
            return today.ToString("M");
        }

        static void fillOrderSheet(Dictionary<string, SalesData>d, string file)
        {
            Excel.Application xl = new Excel.Application();
            Excel.Workbook xlWorkbook = xl.Workbooks.Open(file);
            Excel.Worksheet xlSheet = xlWorkbook.Sheets[1];
            Excel.Range xlRange = xlSheet.UsedRange;
            int rowCount = xlRange.Rows.Count;

            for (int i = 1; i < rowCount; i++)
            {
                string item = (string)(xlRange.Cells[i, 1] as Excel.Range).Value2;
                if (item == null)
                    continue;
                if (d.ContainsKey(item))
                {
                    xlRange.Cells[i, 3].Value2 = d[item].Sold;
                    xlRange.Cells[i, 4].Value2 = d[item].OH;
                }
            }
            xlRange.Cells[1, 5].Value2 = getToday();
            string saveAs = @"C:\Dropbox\Work\Inventory\SP-" + getToday();
            xlWorkbook.SaveAs(saveAs);
            xlWorkbook.Close(file);
            Console.WriteLine("Inventory copy completed");

        }

        static void printMenu()
        {
            Console.WriteLine("**********************************");
            Console.WriteLine(" 1.  Get Inventory");
            Console.WriteLine(" 2.  Generate Perpetual");
            Console.WriteLine(" 3.  Generate Order Sheet");
            Console.WriteLine(" 4.  Quit");
            Console.WriteLine();
            Console.WriteLine(" 5.  Name Conversion (pw required)");
            Console.WriteLine("**********************************");
        }

        static void Main(string[] args)
        {
            Dictionary<string, SalesData> Inventory = new Dictionary<string, SalesData>();
            Inventory.Clear();
            
            int choice = 0;
            while (choice !=4)
            {
                printMenu();
                choice = Int32.Parse(Console.ReadLine());
                switch(choice)
                {
                    case 1:
                        getInventory(Inventory);
                        break;

                    case 2:
                        // Scotch Plains
                        // fillPerpetual(Inventory, @"C:\Dropbox\Work\Inventory\Perpetual-Blank.xlsx");

                        // Chatham
                        fillPerpetual(Inventory, @"C:\Dropbox\Work\Chatham\Perpetual-Blank.xlsx");
                        break;

                    case 3:
                        fillOrderSheet(Inventory, @"C:\Dropbox\Work\Inventory\SPBlank.xls");
                        break;

                    case 4:
                        return;

                    case 5:
                        string pw;
                        pw = Console.ReadLine();
                        if (pw == "DefinitelyNotAdmin")
                            orderConversion(Inventory, @"C:\Dropbox\\Work\Chatham\Perpetual-Blank.xlsx");

                        else
                            Console.WriteLine("Invalid Password.");

                        break;

                    default:
                        Console.WriteLine("**************************");
                        Console.WriteLine("Invalid choice.");
                        break;
                }
            }

            

        }
    }

}
