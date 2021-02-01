using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace WriteLargeExcelFileEfficiently
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var fileName = Path.Combine(path,  @"test.xlsx");
            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }


            var headerList = new string[] { "Header 1", "Header 2", "Header 3", "Header 4" };
            //sheet1 
            var boolList = new bool[] { true, false, true, false };
            var intList = new int[] { 1, 2, 3, -4 };
            var dateList = new DateTime[] { DateTime.Now, DateTime.Today, DateTime.Parse("1/1/2014"), DateTime.Parse("2/2/2014") };
            var sharedStringList = new string[] { "shared string", "shared string", "cell 3", "cell 4" };
            var inlineStringList = new string[] { "inline string", "inline string", "3>", "<4" };


            
            var stopWatch = new Stopwatch();


            using (var spreadSheet = SpreadsheetDocument.Create(fileName, SpreadsheetDocumentType.Workbook))
            {
                // create the workbook
                var workbookPart = spreadSheet.AddWorkbookPart();
                
                var openXmlExportHelper = new OpenXmlWriterHelper();
                openXmlExportHelper.SaveCustomStylesheet(workbookPart);


                var workbook = workbookPart.Workbook = new Workbook();
                var sheets = workbook.AppendChild<Sheets>(new Sheets());



                // create worksheet 1
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet1" };
                sheets.Append(sheet);

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {

                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    //Create header row
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < headerList.Length; i++)
                    {
                        //header formatting attribute.  This will create a <c> element with s=2 as its attribute
                        //s stands for styleindex
                        var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "2") }.ToList();
                        openXmlExportHelper.WriteCellValueSax(writer, headerList[i], CellValues.SharedString, attributes);

                    }
                    writer.WriteEndElement(); //end of Row tag




                    // bool List 
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < boolList.Length; i++)
                    {
                        openXmlExportHelper.WriteCellValueSax(writer, boolList[i].ToString(), CellValues.Boolean);
                    }

                    writer.WriteEndElement(); //end of Row

                    //int List 
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < intList.Length; i++)
                    {
                        openXmlExportHelper.WriteCellValueSax(writer, intList[i].ToString(), CellValues.Number);
                    }
                    writer.WriteEndElement(); // end of Row

                    // datetime List
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < dateList.Length; i++)
                    {
                        //date format.  Excel internally represent the datetime value as number, the date is only a formatting
                        //applied to the number.  It will look something like 40000.2833 without formatting
                        var attributes = new OpenXmlAttribute[] { new OpenXmlAttribute("s", null, "1") }.ToList();
                        //the helper internally translate the CellValues.Date into CellValues.Number before writing
                        openXmlExportHelper.WriteCellValueSax(writer, (dateList[i]).ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Date, attributes);

                    }
                    writer.WriteEndElement(); //end of Row

                    // shared string List
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < sharedStringList.Length; i++)
                    {
                        openXmlExportHelper.WriteCellValueSax(writer, sharedStringList[i], CellValues.SharedString);
                    }
                    writer.WriteEndElement(); //end of Row

                    // inline string List
                    writer.WriteStartElement(new Row());
                    for (int i = 0; i < inlineStringList.Length; i++)
                    {
                        openXmlExportHelper.WriteCellValueSax(writer, inlineStringList[i], CellValues.InlineString);
                    }
                    writer.WriteEndElement(); //end of Row




                    writer.WriteEndElement(); //end of SheetData
                    writer.WriteEndElement(); //end of worksheet
                    writer.Close();
                }










                //sheet2
                Console.WriteLine("Starting to generate 1 million random string");

                //1 million rows
                var a = new string[1000000, 4];
                var chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
                var random = new Random();
                for (int i = 0; i < 1000000; i++)
                {
                    for (int j = 0; j < 4; j++)
                    {
                        a[i, j] = new string(
                            Enumerable.Repeat(chars, 5)
                                        .Select(s => s[random.Next(s.Length)])
                                        .ToArray());
                    }
                }

                Console.WriteLine("Starting to generate 1 million excel rows...");
                
                stopWatch.Start();

                // create worksheet 2
                worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 2, Name = "Sheet2" };
                sheets.Append(sheet);

                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {

                    writer.WriteStartElement(new Worksheet());
                    writer.WriteStartElement(new SheetData());

                    //Create header row

                    for(int i = 0; i < 1000000; i++)
                    {
                        writer.WriteStartElement(new Row());
                        for(int j = 0; j < 4; j++)
                        {
                            openXmlExportHelper.WriteCellValueSax(writer, a[i, j], CellValues.InlineString);
                        }

                        writer.WriteEndElement(); //end of Row tag
                    
                    }



                    writer.WriteEndElement(); //end of SheetData
                    writer.WriteEndElement(); //end of worksheet
                    writer.Close();
                }


                //create the share string part using sax like approach too
                openXmlExportHelper.CreateShareStringPart(workbookPart);
            }
            stopWatch.Stop();

            Console.WriteLine(string.Format("Time elapsed for writing 1 million rows: {0}", stopWatch.Elapsed));
            Console.ReadLine();




        }


    }
}
