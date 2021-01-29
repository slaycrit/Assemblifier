using ClosedXML.Excel;
using Ionic.Zip;
using System;
using System.Collections.Generic;
using System.IO;

namespace Assemblifier
{
    class Program
    {
        //=========================================================================================

        #region Types


        /// <summary>
        /// List of available Manufacturers
        /// </summary>
        enum Manufacturer
        {
            JLCPCB = 0,
        };

        
        #endregion Types

        //=========================================================================================

        #region Fields


        /// <summary>
        /// List of Prefixes (C, IC, Etc.) to Filter Parts for Assembly, use all Parts when this array is empty
        /// </summary>
        static string[] Prefixes;

        /// <summary>
        /// Targeted Service (Board Manufacturer)
        /// </summary>
        static Manufacturer TargetService = Manufacturer.JLCPCB;

        /// <summary>
        /// Output Directory
        /// </summary>
        static string OutputDirectory;


        #endregion Fields

        //=========================================================================================

        #region Methods


        /// <summary>
        /// Main Entry Point
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Console.WriteLine(Environment.CurrentDirectory);

            //Print Arguments
            PrintStringArray(args);

            //Parse Arguments
            foreach(string arg in args)
            {
                //Check for correct Argument Format
                string[] split = arg.Split(':');
                if (split.Length == 2 && split[0].StartsWith('-'))
                {
                    switch (split[0].Substring(1).ToLower())
                    {
                        case "prefix": //Component Name Filter
                            Prefixes = split[1].Split(',');
                            Console.WriteLine($"ARGS: Prefixes Set ({Prefixes})");
                            break;

                        case "service": //Target Manufacturer Service
                            switch (split[1].ToLower())
                            {
                                case "jlcpcb": TargetService = Manufacturer.JLCPCB; break;
                                default: PrintWarning($"Unknown Service \"{split[1]}\"", true); break;
                            }
                            break;

                        default: //Unknown Argument
                            PrintWarning($"Unknown Argument \"{split[0]}\"", true);
                            break;
                    }
                }
                else PrintWarning($"WARNING: Argument Format Fault for \"{arg}\"", true);
            }

            //Set Output File Etc based on Parameters
            OutputDirectory = Environment.CurrentDirectory + "/" + TargetService.ToString() + "/";

            //Find newest Manufacturing Output (Eagle CAM Output), should be inside the Working Directory
            FileInfo archiveFile = null;
            foreach(FileInfo file in new DirectoryInfo(Environment.CurrentDirectory).GetFiles())
            {
                if (file.Extension == ".zip" && (archiveFile == null || file.CreationTime > archiveFile.CreationTime)) archiveFile = file;
            }
            //Fail if no archive was Found
            if(archiveFile == null)
            {
                PrintError("No Zip Archive (CAM Output) Found in the Working Directory");
                return; //EXIT
            }

            //Read Archive Content and extract to Memory Streams (BOM.csv, PnP_back.csv, PnP_front.csv)
            MemoryStream bomStream = null;
            MemoryStream pnpFrontStream = null;
            MemoryStream pnpBackStream = null;
            try
            {
                using (ZipFile zip = ZipFile.Read(archiveFile.Name))
                {
                    foreach(ZipEntry entry in zip)
                    {
                        if(entry.FileName == "BOM.csv")
                        {
                            bomStream = new MemoryStream();
                            entry.Extract(bomStream);
                        }
                        else if(entry.FileName == "PnP_front.csv")
                        {
                            pnpFrontStream = new MemoryStream();
                            entry.Extract(pnpFrontStream);
                        }
                        else if(entry.FileName == "PnP_back.csv")
                        {
                            pnpBackStream = new MemoryStream();
                            entry.Extract(pnpBackStream);
                        }
                    }
                }
            }
            catch (Exception e)
            {
                PrintError($"File IO Error while Reading Archive File \"{archiveFile.Name}\" with error Message\n{e.Message}");
                return; //EXIT
            }

            //Compile BOM Data
            if (bomStream == null) PrintWarning("No BOM Data found in CAM Output, continue to Skip", true);
            else
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    IXLWorksheet ws = wb.AddWorksheet("Sheet1");
                    StreamReader reader = new StreamReader(bomStream);

                    //Write Header
                    ws.Cell("A1").Value = "Comment";
                    ws.Cell("B1").Value = "Designator";
                    ws.Cell("C1").Value = "Footprint";
                    ws.Cell("D1").Value = "LCSC Part #（optional";
                    var rngHeader = ws.Range("A1:D1");
                    rngHeader.Style.Fill.BackgroundColor = XLColor.Yellow;

                    //Write Parts (By Values)

                    //Visual Flair
                    ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Hair;

                    //Save to File
                    try
                    {
                        wb.SaveAs(OutputDirectory + "BOM.xlsx");
                    }
                    catch (Exception e)
                    {
                        PrintError($"File Error while Writing to BOM Output File\n{e.Message}");
                        return; //EXIT
                    }

                    PrintInfo("BOM Output File Saved to Ouput Directory");
                }
            }

            //Compile PnP Data
            if (pnpFrontStream == null && pnpBackStream == null) PrintWarning("No PnP Data found in CAM Output, continue to skip", true);
            else if (pnpFrontStream == null ^ pnpBackStream == null) PrintWarning("Only one PnP Side present, other side will be skipped");
            else
            {
                using (XLWorkbook wb = new XLWorkbook())
                {
                    IXLWorksheet ws = wb.AddWorksheet("Sheet1");
                    StreamReader reader = new StreamReader(bomStream);

                    //Write Header
                    ws.Cell("A1").Value = "Designator";
                    ws.Cell("B1").Value = "Mid X";
                    ws.Cell("C1").Value = "Mid Y";
                    ws.Cell("D1").Value = "Layer";
                    ws.Cell("E1").Value = "Rotation";
                    var rngHeader = ws.Range("A1:E1");
                    rngHeader.Style.Fill.BackgroundColor = XLColor.Yellow;

                    //Write Parts (By Designators)


                    //Visual Flair
                    ws.RangeUsed().Style.Border.OutsideBorder = XLBorderStyleValues.Hair;

                    //Save to File
                    try
                    {
                        wb.SaveAs(OutputDirectory + "CPL.xlsx");
                    }
                    catch (Exception e)
                    {
                        PrintError($"File Error while Writing to CPL Output File\n{e.Message}");
                        return; //EXIT
                    }

                    PrintInfo("CPL Output File Saved to Ouput Directory");
                }
            }

            Console.ReadLine();
        }

        /// <summary>
        /// Helper Function for Printing string Arrays like the Arguments
        /// </summary>
        static void PrintStringArray(string[] arr)
        {
            int i = 0;
            foreach (string element in arr)
            {
                Console.WriteLine($"{i++,2}: \"{element}\"");
            }
        }

        /// <summary>
        /// Prints a simple Informational Message to the Console
        /// </summary>
        /// <param name="message">Message Printed on the Console</param>
        static void PrintInfo(string message)
        {
            Console.WriteLine($"INFO: {message}");
        }

        /// <summary>
        /// Prints an Warning Message to the Console and waits for the User to choose an option if enabled
        /// </summary>
        /// <param name="message">Message Printed on the Console</param>
        /// <param name="confirm">Asks User to press 'y' to continue or 'n' to abort when True</param>
        /// <returns>True when Continuation is desired</returns>
        static bool PrintWarning(string message, bool confirm = false)
        {
            Console.WriteLine($"WARNING: {message}");
            if (confirm)
            {
                Console.WriteLine("Continue? y/n: ");
                while (true)
                {
                    ConsoleKeyInfo consoleKey = Console.ReadKey();
                    if (consoleKey.KeyChar == 'y') return true;
                    else if (consoleKey.KeyChar == 'n') return false;
                }
            }
            else return true;
        }

        /// <summary>
        /// Prints an Error Message to the Console and waits for the User to acknowledge it before returning
        /// </summary>
        /// <param name="message">Message Printed on the Console</param>
        static void PrintError(string message)
        {
            Console.WriteLine($"ERROR: {message}");
            Console.ReadLine();
        }

        #endregion
    }
}
