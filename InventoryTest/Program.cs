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
        };

        static void getInventory(Dictionary<string, SalesData> d)
        {
            char[] delimiterChars = { ',' };
            string[] productCodes = { "584-10", "584-30", "555-10", "580-00", "581-00", "587-00" };
            string line;
            StreamReader file = new StreamReader("Report.txt");

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
                d.Add(words[2], s);
            }
            file.Close();
        }

        static void xlTest()
        {
            Excel.Application xcel = new Excel.Application();
            Excel.Workbook wkb = xcel.Workbooks.Open(@"D:\InventoryTest\InventoryTest\bin\Debug\Perpetual-Blank.xlsx");
            Excel.Worksheet sheet = wkb.Sheets[1];
            Excel.Range xlrange = sheet.UsedRange;
            int rowCount = xlrange.Rows.Count;
            int colCount = xlrange.Columns.Count;

            for (int i = 1; i < rowCount; i++)
            {
                string item;
                item = (string)(xlrange.Cells[i, 1] as Excel.Range).Value2;
                Console.WriteLine(item);
            }


            wkb.Close();
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
            string saveAs = @"D:\InventoryTest\InventoryTest\bin\Debug\Perpetual-WE-11-27.xlsx";
            
            xlWorkbook.SaveAs(saveAs);
            xlWorkbook.Close(file);

            Console.WriteLine("Copy completed.");

        }

        static void Main(string[] args)
        {
            Dictionary<string, SalesData> Inventory = new Dictionary<string, SalesData>();
            Inventory.Clear();
            getInventory(Inventory);
            //xlTest();

            //  Set file equal to the location of the blank perpetual file.
            string file = @"D:\InventoryTest\InventoryTest\bin\Debug\Perpetual-Blank.xlsx";
            fillPerpetual(Inventory, file);

            //nameConversion(Inventory, file);
            Console.ReadLine();
        }
    }

}
