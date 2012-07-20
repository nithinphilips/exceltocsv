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
            TryKillProcessByMainWindowHwnd(hwnd);

            return 0;
        }

        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary> Tries to find and kill process by hWnd to the main window of the process.</summary>
        /// <param name="hWnd">Handle to the main window of the process.</param>
        /// <returns>True if process was found and killed. False if process was not found by hWnd or if it could not be killed.</returns>
        public static bool TryKillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            if (processID == 0) return false;
            try
            {
                Process.GetProcessById((int)processID).Kill();
            }
            catch (ArgumentException)
            {
                return false;
            }
            catch (Win32Exception)
            {
                return false;
            }
            catch (NotSupportedException)
            {
                return false;
            }
            catch (InvalidOperationException)
            {
                return false;
            }
            return true;
        }

        /// <summary> Finds and kills process by hWnd to the main window of the process.</summary>
        /// <param name="hWnd">Handle to the main window of the process.</param>
        /// <exception cref="ArgumentException">
        /// Thrown when process is not found by the hWnd parameter (the process is not running). 
        /// The identifier of the process might be expired.
        /// </exception>
        /// <exception cref="Win32Exception">See Process.Kill() exceptions documentation.</exception>
        /// <exception cref="NotSupportedException">See Process.Kill() exceptions documentation.</exception>
        /// <exception cref="InvalidOperationException">See Process.Kill() exceptions documentation.</exception>
        public static void KillProcessByMainWindowHwnd(int hWnd)
        {
            uint processID;
            GetWindowThreadProcessId((IntPtr)hWnd, out processID);
            if (processID == 0)
                throw new ArgumentException("Process has not been found by the given main window handle.", "hWnd");
            Process.GetProcessById((int)processID).Kill();
        }
    }
}
