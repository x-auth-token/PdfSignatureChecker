using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;

using iTextSharp.text.pdf;
using iTextSharp.text.exceptions;

namespace SignatureCheckerV2
{
    /*
        TO DO:
            1. Check InvalidPdfException 
            2. Check of files is true pdf
            3. Resolve AGPL License issue
    */


    class Program
    {
        private static IEnumerable<string> filenames;
        private static string srcFolder;
        private static string dstFolder;
        private static string currentDate = DateTime.Now.ToString("yyyy-MM-dd-HH-mm-ss");
        private static int row = 1;
        private static Application excel;
        private static Workbook wb;
        private static Worksheet ws;
        private static string resultFile, filename;

        private static void DisplayHelp()
        {
            Console.WriteLine("Program Usage:");
            Console.WriteLine("SignatureCheckerV2 <Source folder> <Output file save path>");
            Console.WriteLine("\t-p --prompt\tPrompt mode. You will provide source and destination path through prompt");
            Console.WriteLine("\t-h --help\tShow this help");

        }

        private static void Menu(String[] args)
        {
            if (args.Length == 0)
            {
                DisplayHelp();
                System.Environment.Exit(0);

            }
            else if (args.Length > 2)
            {
                Console.WriteLine("Invalid number of arguments!");
                System.Environment.Exit(1);
            }
            else if (args.Length == 2)
            {
                //if (args[0].ToLower().Equals("-s") || args[2].ToLower().Equals("-s"))
                srcFolder = args[0];
                dstFolder = args[1];
            }
            else if (args.Length == 1)
            {
                if (args[0].ToLower().Equals("-h") || args[0].ToLower().Equals("--help"))
                {
                    DisplayHelp();
                }
                else if (args[0].ToLower().Equals("-p") || args[0].ToLower().Equals("--prompt"))
                {

                    Console.Write("Enter source folder path. Type \"Exit\" to exit the program: ");
                    do
                    {
                        srcFolder = Console.ReadLine();

                        if (srcFolder.ToLower().Equals("exit"))
                        {
                            Console.WriteLine("Exiting");
                            System.Environment.Exit(0);
                        }

                        if (srcFolder.Length == 0)
                        {
                            Console.WriteLine("Source Folder Path cannot be empty!");
                            Console.Write("Enter source folder path. Type \"Exit\" to exit the program: ");
                        }

                    } while (srcFolder.Length == 0);
                }


                Console.Write("Destination Directory (Press Enter to create the file in default location):");
                dstFolder = Console.ReadLine();

                if (dstFolder.Length == 0)
                {
                    dstFolder = @"C:\Temp";
                    Console.WriteLine("The file will be created in \"C:\\Temp\"");
                }



            }
            else
            {
                Console.WriteLine("Invalid argument!");
                System.Environment.Exit(1);
            }
        }
        static void Main(string[] args)
        {

            //string resultFile = Directory.GetCurrentDirectory().ToString() + @"\Result_" + currentDate + ".xlsx";
            try
            {

                Menu(args);

                if (!Directory.Exists(dstFolder))
                {
                    Directory.CreateDirectory(dstFolder);
                }

                resultFile = dstFolder + @"\Result_" + currentDate + ".xlsx";

                // Creating excel file
                excel = new Application { Visible = false, DisplayAlerts = false };
                wb = excel.Workbooks.Add(Type.Missing);
                ws = wb.ActiveSheet;
                ws.Name = "Result";

                // Set first row headers to filter later
                ws.Cells[row, 1] = "Filename";
                ws.Cells[row, 2] = "Signing Status";
                row++;


                // Get all target filenames from directory including sub-directories               
                filenames = Directory.EnumerateFiles(srcFolder, "*.pdf", SearchOption.AllDirectories);

                foreach (string temp in filenames)
                {
                    filename = temp;
                    try
                    {
                        using (var doc = new PdfReader(filename))
                        {

                            AcroFields acroFields = doc.AcroFields;

                            // Checking if signature fields exist
                            if (acroFields.GetSignatureNames().Count != 0)
                            {


                                ws.Cells[row, 1] = filename;
                                ws.Cells[row, 2] = "signed";
                                row++;
                                //wb.SaveAs(resultFile);
                            }
                            else
                            {
                                ws.Cells[row, 1] = filename;
                                ws.Cells[row, 2] = "unsigned";
                                row++;
                                //wb.SaveAs(resultFile);

                            }

                        }
                    }
                    catch (InvalidPdfException e)
                    {

                        ws.Cells[row, 1] = filename;
                        ws.Cells[row, 2] = "Corrupted or non conformant to PDF standard";
                        row++;
                    }
                }

                // Adding Autofilter
                
                /*ws.UsedRange.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);
                ws.Columns.AutoFit();
                wb.SaveAs(resultFile);

                excel.Visible = true;*/

                

            } // end try
            catch (DirectoryNotFoundException e)
            {
                Console.WriteLine(e.Message);
            }

            

            ws.UsedRange.AutoFilter(1, Type.Missing, XlAutoFilterOperator.xlAnd, Type.Missing, true);
            ws.Columns.AutoFit();
            wb.SaveAs(resultFile);

            excel.Visible = true;
        }


    }
}
