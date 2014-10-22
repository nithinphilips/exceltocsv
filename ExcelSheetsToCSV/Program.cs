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
            // The variable are to be set via command-line options
            bool help = false;
            bool printFormats = false;
            bool stdout = false;
            int sheetIndex = -1;
            string outputType = "csv";
            string outFileName = string.Empty;
            string workSheetPattern = string.Empty;
            XlFileFormat format = XlFileFormat.xlCSVWindows;
            bool noprompt = false;

            // The command-line option specification
            var p = new OptionSet()
            {
                { "f|format=", v => outputType = v},
                { "o|output=", v => outFileName = v},
                { "h|help", v => help = v  != null},
                { "listformats", v => printFormats = v  != null},
                { "stdout", v => stdout = v  != null},
                { "y|force", v => noprompt = v  != null},
                { "i|index=", v => sheetIndex = int.Parse(v)},
                { "g|pattern=", v => workSheetPattern = v},
            };


            // the parse commandline options and get any arguments
            List<string> files = p.Parse(args);

            // We support some shortcuts for setting output formats.
            // Check if the user specified a shortcut, otherwise try to parse into an Excel enum.
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

            // User wants to see a list of all supported formats
            if (printFormats)
            {
                Console.WriteLine("Format can be any of these values:\n");
                var values = Enum.GetValues(typeof(XlFileFormat));
                foreach (XlFileFormat val in values)
                {
                    Console.WriteLine("\t{0}", val);
                }
                return 0;
            }

            // Specified help option or omitted required arguments?
            if (help || files.Count <= 0)
            {
                Console.WriteLine("Excel Sheets to CSV Converter");
                Console.WriteLine();
                Console.WriteLine("Usage:");
                Console.WriteLine("\texcelsheetstocsv.exe --help | --listformats | ([--output=<outputfilename>] [--index=<index>] [--pattern=<regex-patten>] [--format=csv|tab]  <excel-file> [<excel-file> ...])");
                Console.WriteLine();

                Console.WriteLine("If both --pattern and --index is given, --index wins.");
                Console.WriteLine("Use the '--listformats' option to get a full listing of possible format values");
                Console.WriteLine("'--ouputfilename' can be used with '--index' to save a single worksheet to a specific file.");
                Console.WriteLine("");

                return 0;
            }

            // Expand any wild card filenames.
            // The support is rudimentary, only files in working directory are searched and only wilcards supported by windows work.
            // Example: *.xlsx
            var newFilesList = new List<string>();
            foreach (var file in files)
            {
                if (file.StartsWith("*"))
                    newFilesList.AddRange(Directory.GetFiles(Environment.CurrentDirectory, file));
                else
                    newFilesList.Add(file);
            }

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlApp.DisplayAlerts = false;
            
            
            // Transform each file
            foreach (var file in newFilesList)
            {
                if (stdout)
                {
                    outFileName = Path.GetTempFileName();
                    File.Delete(outFileName);
                }

                var fi = new FileInfo(file);
                if (fi.Exists)
                {
                    if (string.IsNullOrWhiteSpace(outFileName))
                    {
                        outFileName = Path.GetFileNameWithoutExtension(fi.Name);
                    }


                    if (sheetIndex != -1)
                    {
                        SaveWorkSheet(GetWorkSheets(fi.FullName)[sheetIndex], outFileName, format, true, !noprompt);

                        if (stdout)
                        {
                            using (var stream = new FileStream(outFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                            {
                                using (var textReader = new StreamReader(stream))
                                {
                                    Console.WriteLine(textReader.ReadToEnd());
                                }
                            }
                        }
                    }
                    else if (!string.IsNullOrWhiteSpace(workSheetPattern))
                    {
                        Console.Error.WriteLine("Using pattern matching...");
                        Worksheet firstWorksheet = GetWorkSheets(fi.FullName, workSheetPattern).FirstOrDefault();

                        if (firstWorksheet != null)
                        {
                            Console.Error.WriteLine("Extracting the first matching worksheet named {0}...", firstWorksheet.Name);
                            SaveWorkSheet(firstWorksheet, outFileName, format, true, !noprompt);

                            if (stdout)
                            {
                                using (var stream = new FileStream(outFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                                {
                                    using (var textReader = new StreamReader(stream))
                                    {
                                        Console.WriteLine(textReader.ReadToEnd());
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.Error.WriteLine("No worksheets matched the pattern '{}'", workSheetPattern);
                            return 102;
                        }
                    }
                    else
                    {
                        TransformSpreadSheet(fi.FullName, format, !noprompt);
                    }
                }
                else
                {
                    Console.Error.WriteLine("An input file you specified, '{0}', does not exist", fi.Name);
                }
            }

            int hwnd = xlApp.Application.Hwnd;
            xlApp.Quit();
            Helper.TryKillProcessByMainWindowHwnd(hwnd);
            
            // done!
            return 0;
        }

        static Microsoft.Office.Interop.Excel.Application xlApp;


        /// <summary>
        /// 
        /// </summary>
        /// <param name="excelFile"></param>
        /// <param name="searchPattern"></param>
        /// <returns></returns>
        static IEnumerable<Worksheet> GetWorkSheets(string excelFile, string searchPattern)
        {
            System.Text.RegularExpressions.Regex r = new System.Text.RegularExpressions.Regex(searchPattern, System.Text.RegularExpressions.RegexOptions.Singleline | System.Text.RegularExpressions.RegexOptions.IgnoreCase);

            foreach (var workSheet in GetWorkSheets(excelFile))
            {
                var sheet = workSheet as Worksheet;

                if (r.IsMatch(sheet.Name))
                {
                    yield return sheet;
                }
            }
        }

        static Sheets GetWorkSheets(string excelFile)
        {
            Console.Error.WriteLine("Reading {0}...", excelFile);

            Workbook book = xlApp.Workbooks.Open(excelFile);

            return book.Worksheets;
        }

        static void TransformSpreadSheet(string excelFile, XlFileFormat format, bool prompt)
        {
            foreach (var workSheet in GetWorkSheets(excelFile))
            {
                var sheet = workSheet as Worksheet;

                var csvFileName = Path.GetFileNameWithoutExtension(excelFile) + "-" + sheet.Name + ".csv";

                SaveWorkSheet(sheet, csvFileName, format, false, prompt);
            }
        }

        static void SaveWorkSheet(Worksheet sheet, string csvFileName, XlFileFormat format, bool quiet, bool prompt)
        {
            if(!quiet) Console.Error.WriteLine("Saving worksheet {0} to file '{1}'", sheet.Name, csvFileName);
            csvFileName = Path.Combine(Environment.CurrentDirectory, csvFileName);

            if (File.Exists(csvFileName))
            {
                if (prompt)
                {
                    Console.Error.Write("The file {0} exists. Replace (Y/n)?", Path.GetFileName(csvFileName));
                    var input = Console.ReadLine();
                    if (input == "n") return;
                }

                File.Delete(csvFileName);
            }
            sheet.SaveAs(csvFileName, format);
        }

    }
}
