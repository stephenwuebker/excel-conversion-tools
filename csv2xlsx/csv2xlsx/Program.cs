using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace csv2xlsx
{
    class Program
    {
        static void Main(string[] args)
        {
            string csvPath = "";
            string newFilename = "";

            // make sure the correct arguments were passed
            if (args.Count() > 2 || args.Count() < 1)
            {
                Console.WriteLine("csv2xlsx takes one or two parameters. Call csv2xlsx with the file you wish to convert and optionally where you want to save it.");
                return;
            }

            csvPath = args[0].ToString();

            if (args.Count() == 2)
                newFilename = args[1].ToString();

            // this might not be required. excel should be able to open anything -- maybe
            //// make sure it's excel format
            //if (!xlsPath.Substring(xlsPath.Length-3, 3).Equals("xls") || !xlsPath.Substring(xlsPath.Length-4, 4).Equals("xlsx"))
            //{
            //    Console.WriteLine("Invalid file type. Please enter an excel file type (.xls or .xlsx)");
            //    return;
            //}

            SaveAs(csvPath, newFilename);


        }


        public static void SaveAs(string csvPath, string newFilename)
        {
            try
            {
                Application app = new Application();
                Workbook wbWorkbook = app.Workbooks.Open(csvPath); //app.Workbooks.Add(Type.Missing);


                if (newFilename == "")
                    newFilename = csvPath.Substring(0, csvPath.LastIndexOf(".")) + ".xslx";

                wbWorkbook.SaveAs(newFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlOpenXMLWorkbook,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                wbWorkbook.Close(false, csvPath, true);
                app.Quit();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
