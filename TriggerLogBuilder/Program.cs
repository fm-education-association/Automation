using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using XL = Microsoft.Office.Interop.Excel;

namespace TriggerLogBuilder
{
    class Program
    {
        static void Main(string[] args)
        {
            DoWork();
            Console.WriteLine("Done, press enter");
            Console.ReadLine();

        }
        static void DoWork()
        {
            var rowCount = 50;
            var colCount = 7;
            var items = new Dictionary<string, List<string>>();
            foreach (var key in new string[] { "head", "hour", "foot", "type" })
            {
                items.Add(key, new List<string>());
            }

            //Create COM Objects. Create a COM object for everything that is referenced
            XL.Application xlApp; // = new Excel.Application();
            xlApp = (XL.Application)Marshal.GetActiveObject("Excel.Application");


            if (xlApp is null || !(xlApp is XL.Application) || xlApp.ActiveWorkbook is null)
            {
                Console.WriteLine("COULD NOT FIND EXCEL. MUST BE OPEN WITH TRIGGER-BASED SCHEDULE IMPORT _TEMPLATE_ OPEN IN IT.");
                return;
            }
            // Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"sandbox_test.xlsx");
            XL.Workbook xlWorkbook = xlApp.ActiveWorkbook;
            XL._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
            XL.Range xlRange = xlWorksheet.UsedRange;
            try
            {
                // Then you can read from the sheet, keeping in mind that indexing in Excel is not 0 based.This just reads the cells and prints them back just as they were in the file.

                //iterate over the rows and columns and print to the console as it appears in the file
                //excel is not zero based!!



                for (int i = 1; i <= rowCount; i++)
                {
                    Console.Write($"Reading line {i}\r      ");
                    var LineString = "";
                    var LineCategory = "";
                    for (int j = 1; j <= colCount; j++)
                    {
                        var CellVal = "";
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            CellVal = xlRange.Cells[i, j].Value2.ToString();
                        }
                        if (j == 1) LineCategory = CellVal.ToLower();
                        if (j > 2) LineString += "|";
                        if (j >= 2) LineString += CellVal;
                    }
                    if (i == 1)
                    {
                        if (LineCategory != "section" || LineString != "Cue|Sched|Name|Length|Category|Description")
                        {
                            Console.WriteLine("Wrong spreadsheet, requires headings like Section,Cue,Sched,Name,Length,Category,Description (you are going to want a sample!!)");
                            return;
                        }
                    }
                    else
                    {
                        if (LineCategory.Trim() != "")
                        {
                            List<string> lst = null;
                            if (!items.TryGetValue(LineCategory, out lst))
                            {
                                Console.WriteLine($"Line {i} Category (col 1 - '{LineCategory}') not in \"head\",\"hour\",\"foot\",\"type\"");
                                return;
                            }
                            lst.Add(LineString);
                        }

                    }
                }
                Console.WriteLine("I've Read all lines        ");
                // Lastly, the references to the unmanaged memory must be released.If this is not properly done, then there will be lingering processes that will hold the file access writes to your Excel workbook.

            }
            finally
            {
                //cleanup
                GC.Collect();
                GC.WaitForPendingFinalizers();

                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);
                Marshal.ReleaseComObject(xlWorkbook);
                Marshal.ReleaseComObject(xlApp);
            }


            // this will hold output
            var outFile = new System.Text.StringBuilder();

            foreach (var item in items["head"])
            {
                outFile.AppendLine(item);
            }

            for (var hour = 0; hour <= 23; hour++)
            {
                var hourtext = string.Join("\r\n", items["hour"]);
                var defaultcarts = from item in items["type"] where item.Split('|')[1].ToLower() == "default" select item;
                if (!defaultcarts.Any() || defaultcarts.First().Split('|')[2].Trim() == "")
                {
                    Console.WriteLine("There has to be a line of category type with a default hour (sched) and cart name (name)");
                    return;
                }

                if (defaultcarts.Count() > 1)
                {
                    Console.WriteLine($"More than one type entry for an hour, I can only handle one - {defaultcarts.Skip(1).First()}");
                    return;
                }

                var Cart4Hour = defaultcarts.First().Split('|')[2].Trim();


                var hourcarts = from item in items["type"]
                                where item.Split('|')[1].Trim() == hour.ToString()
                                   || item.Split('|')[1].Trim() == ("0" + hour.ToString())
                                select item;
                if (hourcarts.Any())
                {
                    if (hourcarts.Count() > 1)
                    {
                        Console.WriteLine($"More than one type entry for an hour, I can only handle one - {hourcarts.Skip(1).First()}");
                        return;
                    }

                    var typecart = hourcarts.First().Split('|')[2].Trim();
                    if (typecart == "")
                    {
                        Console.WriteLine($"Blank cart name for type with hour listed {hourcarts.First()}");
                        return;
                    }
                    Cart4Hour = typecart;
                }
                Console.WriteLine($"Hour {hour} cart is {Cart4Hour}");
                outFile.AppendLine(hourtext.Replace("{hour}", hour.ToString("D2")).Replace("{type}", Cart4Hour));
            }

            foreach (var item in items["foot"])
            {
                outFile.AppendLine(item);
            }
            System.IO.File.WriteAllText("TriggeredLog.txt", outFile.ToString());
          }
    }
}
