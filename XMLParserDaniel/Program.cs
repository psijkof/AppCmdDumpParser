using System;
using System.IO;
using System.Xml;
using System.Xml.Linq;
using System.Linq;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;

namespace XMLParserDaniel
{
    class Program
    {
        static void Main(string[] args)
        {
            var appCmdDump = XElement.Load($@"{Path.GetTempPath()}\RGPD-sites-aspserver-compressed.xml");
            var outputSheet = Path.Combine(Path.GetTempPath(), "sitesExcel.xlsx");

            var sites = appCmdDump.Elements("SITE");
            var siteName = "";
            XElement physicalPath;
            var physPathName = "";
            var bindingInfo = "";


            using (var fs = new FileStream(outputSheet, FileMode.Create, FileAccess.Write))
            {
                // create xlsx file
                var workbook = new XSSFWorkbook();
                // write first row with header info
                var sheet = CreateHeader(workbook);

                var previousSiteName = "";
                var rowNr = 0;

                foreach (var el in sites)
                {
                    rowNr++;

                    siteName = el.Attribute("SITE.NAME").Value;

                    physicalPath = el.Descendants("virtualDirectory")
                        .Where(e => (string)e.Attribute("path") == "/")
                        .FirstOrDefault();

                    if (physicalPath != null)
                    {
                        physPathName = physicalPath.Attribute("physicalPath").Value;
                    }
                    else
                    {
                        physPathName = "ERROR - no physical path specified!";
                    }

                    var bindings = el.Descendants("binding");

                    foreach (var binding in bindings)
                    {
                        // create a new row per binding info found
                        var row = sheet.CreateRow(rowNr);

                        // only write sitename once
                        row.CreateCell(0, CellType.String).SetCellValue(previousSiteName != siteName ? siteName : "");
                        previousSiteName = siteName;

                        row.CreateCell(1, CellType.String).SetCellValue(physPathName);

                        // create consumable url from binding info
                        bindingInfo = binding.Attribute("bindingInformation").Value;
                        var protocol = binding.Attribute("protocol").Value;
                        var biParts = bindingInfo.Split(':');
                        var url = protocol + "://" + biParts[biParts.Length - 1] + ":" + biParts[1];

                        // write binding info
                        row.CreateCell(2, CellType.String).SetCellValue(bindingInfo);
                        row.CreateCell(3, CellType.String).SetCellValue(url);
                        //row.GetCell(3).Hyperlink = new XSSFHyperlink(HyperlinkType.Url);
                        //try
                        //{
                        //    row.GetCell(3).Hyperlink.Address = url;
                        //}
                        //catch (Exception)
                        //{
                        //}
                        rowNr++;
                    }
                }

                // resize columns
                sheet.AutoSizeColumn(0);
                sheet.AutoSizeColumn(1);
                sheet.AutoSizeColumn(2);
                sheet.AutoSizeColumn(3);

                // write xlsx file
                workbook.Write(fs);
            }
            Console.WriteLine("Done!");
            Console.ReadLine();
        }

        private static ISheet CreateHeader(XSSFWorkbook workbook)
        {
            var sheet1 = workbook.CreateSheet("sites");
            var row = sheet1.CreateRow(0);
            row.CreateCell(0, CellType.String).SetCellValue("Site Code");
            row.CreateCell(1, CellType.String).SetCellValue("Physical path");
            row.CreateCell(2, CellType.String).SetCellValue("Binding");
            row.CreateCell(3, CellType.String).SetCellValue("Consumable Url");

            return sheet1;
        }
    }
}
