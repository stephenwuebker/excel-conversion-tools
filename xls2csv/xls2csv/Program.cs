using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;

namespace xls2csv
{
    class Program
    {
        static void Main(string[] args)
        {
            string xlsPath = "";
            string newFilename = "";
            string passWord = "";

            // make sure the correct arguments were passed
            if (args.Count() > 3 || args.Count() < 1)
            {
                Console.WriteLine("xls2csv takes one or two or three parameters. Call xls2csv with the file you wish to convert and optionally where you want to save it. If there is a password on the file, it is the third argument.");
                return;
            }

            xlsPath = args[0].ToString();

            if (args.Count() > 1)
                newFilename = args[1].ToString();

            if (args.Count() > 2)
                passWord = args[2].ToString();

            // this might not be required. excel should be able to open anything -- maybe
            //// make sure it's excel format
            //if (!xlsPath.Substring(xlsPath.Length-3, 3).Equals("xls") || !xlsPath.Substring(xlsPath.Length-4, 4).Equals("xlsx"))
            //{
            //    Console.WriteLine("Invalid file type. Please enter an excel file type (.xls or .xlsx)");
            //    return;
            //}

            SaveAs(xlsPath, newFilename, passWord);


        }


        public static void SaveAs(string xlsPath, string newFilename, string passWord)
        {
            try
            {
                Application app = new Application();
                Workbook wbWorkbook = app.Workbooks.Open(xlsPath, Type.Missing, Type.Missing, Type.Missing, passWord); //app.Workbooks.Add(Type.Missing);


                if(newFilename == "")
                    newFilename = xlsPath.Substring(0, xlsPath.LastIndexOf(".")) + ".csv";

                wbWorkbook.SaveAs(newFilename, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows, 
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, 
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                wbWorkbook.Close(false, xlsPath, true);
                app.Quit();

            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}
