using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.ComponentModel;

namespace ExcelSheetsToCSV
{
    class Program
    {
        static int Main(string[] args)
        {
            if (args.Length != 1)
            {
                Console.WriteLine("Excel Sheets to CSV Converter");
                Console.WriteLine();
                Console.WriteLine("Usage:");
                Console.WriteLine("\texcelsheetstocsv.exe <excel-file>");
                return 0;
            }

            string excelFile = new FileInfo(args[0]).FullName;

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
                sheet.SaveAs(csvFileName, XlFileFormat.xlCSVWindows);
            }

            int hwnd = xlApp.Application.Hwnd;
            xlApp.Quit();
            Helper.TryKillProcessByMainWindowHwnd(hwnd);

            return 0;
        }

    }
}
