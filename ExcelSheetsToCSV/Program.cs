using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.ComponentModel;
using NDesk.Options;

namespace ExcelSheetsToCSV
{
    class Program
    {
        static int Main(string[] args)
        {
            // the options to be set via command-line
            bool help = false;
            string outputType = "csv";
            XlFileFormat format = XlFileFormat.xlCSVWindows;

            // the command-line options
            var p = new OptionSet()
            {
                { "t|type=", v => outputType = v},
                { "h|help", v => help = v  != null},
            };

            // the parse commandline options and get any arguments
            List<string> files = p.Parse(args);

            switch (outputType.ToLower())
            {
                case "csv":
                    format = XlFileFormat.xlCSVWindows;
                    break;
                case "tab":
                    format = XlFileFormat.xlTextWindows;
                    break;
                default:
                    if (!Enum.TryParse(outputType, out format))
                    {
                        Console.WriteLine("I did not recognize the output format you specified. Try again.");
                        return 0;
                    }
                    break;
            }


            // User need help or omitted required argument?
            if (help || files.Count <= 0)
            {
                Console.WriteLine("Excel Sheets to CSV Converter");
                Console.WriteLine();
                Console.WriteLine("Usage:");
                Console.WriteLine("\texcelsheetstocsv.exe --help | ([--type=csv|tab]  <excel-file> [<excel-file> ...])");
                return 0;
            }

            // Transform each file
            foreach (var file in files)
            {
                var fi = new FileInfo(file);
                if (fi.Exists)
                {
                    TransformSpreadSheet(fi.FullName, format);
                }
                else
                {
                    Console.Error.WriteLine("An input file you specified, '{0}', does not exist", fi.Name);
                }
            }
            

            // done!
            return 0;
        }

        static void TransformSpreadSheet(string excelFile, XlFileFormat format)
        {
            Console.WriteLine("Reading {0}...", excelFile);

            Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;

            Workbook book = xlApp.Workbooks.Open(excelFile);

            foreach (var workSheet in book.Worksheets)
            {
                var sheet = workSheet as Worksheet;

                var csvFileName = Path.GetFileNameWithoutExtension(excelFile) + "-" + sheet.Name + ".csv";
                Console.WriteLine("Saving worksheet {0} to file '{1}'", sheet.Name, csvFileName);
                csvFileName = Path.Combine(Environment.CurrentDirectory, csvFileName);

                if (File.Exists(csvFileName))
                {
                    Console.Write("The file {0} exists. Replace (Y/n)?", Path.GetFileName(csvFileName));
                    var input = Console.ReadLine();
                    if (input == "n") continue;

                    File.Delete(csvFileName);
                }
                sheet.SaveAs(csvFileName, format);
            }

            int hwnd = xlApp.Application.Hwnd;
            xlApp.Quit();
            Helper.TryKillProcessByMainWindowHwnd(hwnd);
        }

    }
}
