using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using ExcelApp = Microsoft.Office.Interop.Excel;


namespace AutoMailStudentCheckdown
{
    class Program
    {
        public static void Main(string[] args)
        {
            // creating object of class program
            Program p = new Program();

            Console.WriteLine("=============================================================================================================");
            Console.WriteLine("                             Daytona State: Auto Mailer for Student Checkdowns                               ");
            Console.WriteLine("=============================================================================================================");
            string userPath = Environment.GetEnvironmentVariable("userprofile"); // User path variable if needed
            // Console.WriteLine(userPath);

            //create a file array that searches for all files in specific directory and subdirectories
            Console.Write("Input the path to the student check down files: ");
            string filePathForCheckdowns = Console.ReadLine();
            string[] fileArray = Directory.GetFiles(@filePathForCheckdowns, "*.xlsx", SearchOption.AllDirectories);


            Console.WriteLine("[{0}]", string.Join(", \n", fileArray));

            foreach (string file in fileArray)
            {
                // calling method
                p.ExecuteExcelMacro(file);
            }

            Console.WriteLine("=============================================================================================================");
            Console.WriteLine("                             Press any key to exit the program...                                            ");
            Console.WriteLine("=============================================================================================================");
            Console.ReadKey();
        }

        public void ExecuteExcelMacro(string sourceFile)
        {
            ExcelApp.Application ExcelApp = new ExcelApp.Application();
            ExcelApp.DisplayAlerts = false;
            ExcelApp.Visible = false;
            ExcelApp.Workbook ExcelWorkBook = ExcelApp.Workbooks.Open(sourceFile);
            ExcelApp.Run("BuiltMacro");
            ExcelWorkBook.Close(false);
            ExcelApp.Quit();
            if (ExcelWorkBook != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWorkBook); }
            if (ExcelApp != null) { System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelApp); }
        }
    }
}
