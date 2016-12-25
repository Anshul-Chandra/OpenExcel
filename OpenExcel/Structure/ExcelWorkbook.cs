using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OpenExcel
{
    /// <summary>
    /// Class for Excel Workbook (Consists of Shared Strings too)
    /// </summary>
    public class ExcelWorkbook
    {
        public string FileName { get; set; }
        private Package ExcelFile = null;

        protected XmlDocument workbookDoc = null;
        private XmlElement sheets { get; set; }

        private XmlDocument sharedStringsDoc = null;
        private XmlElement sstNode { get; set; }

        internal int ssCount { get; set; }

        public List<ExcelSheet> Sheets { get; set; }

        #region Constructors

        /// <summary>
        /// Default constructor
        /// </summary>
        public ExcelWorkbook()
        {
            this.FileName = "ExcelWorkbook";

            InitializeWorkbook();
            InitializeSharedStringDoc();
        }

        //Creates an Excel Workbook by the given name
        public ExcelWorkbook(string FileName)
        {
            this.FileName = FileName;

            InitializeWorkbook();
            InitializeSharedStringDoc();
        }

        #endregion

        #region Methods

        private void InitializeWorkbook() //Add a parameter for the location in which the file is saved
        {
            Sheets = new List<ExcelSheet>();
            workbookDoc = new XmlDocument();

            //Obtain a reference to the root node, and then add the XML declaration.
            XmlElement wbRoot = workbookDoc.DocumentElement;
            XmlDeclaration wbxmldecl = workbookDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            workbookDoc.InsertBefore(wbxmldecl, wbRoot);

            //Create and append the workbook node to the document.
            XmlElement workBook = workbookDoc.CreateElement("workbook");
            workBook.SetAttribute("xmlns", "http://schemas.openxmlformats.org/" + "spreadsheetml/2006/main");
            workBook.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/officeDocument/" + "2006/relationships");
            workbookDoc.AppendChild(workBook);

            ////Create and append the workbook node to the document.
            //XmlElement workbookPr = workbookDoc.CreateElement("workbookPr");
            //workbookPr.SetAttribute("defaultThemeVersion", "124226");
            //workBook.AppendChild(workbookPr);

            //// Creating a sheetProctection part for password protection
            //XmlElement workbookProtection = workbookDoc.CreateElement("workbookProtection");
            //workbookProtection.SetAttribute("workbookPassword", "xsd:hexBinary data");
            //workbookProtection.SetAttribute("lockStructure", "1");
            //workbookProtection.SetAttribute("lockWindows", "1");
            //workBook.AppendChild(workbookProtection);

            ////Create and append the workbook node to the document.
            //XmlElement bookViews = workbookDoc.CreateElement("bookViews");
            
            //XmlElement workbookView = workbookDoc.CreateElement("workbookView");
            //workbookView.SetAttribute("windowHeight", "7365");
            //workbookView.SetAttribute("windowWidth", "19815");
            //workbookView.SetAttribute("yWindow", "555");
            //workbookView.SetAttribute("xWindow", "390");
            //bookViews.AppendChild(workbookView);
            
            //workBook.AppendChild(bookViews);

            //Create and append the sheets node to the workBook node.
            sheets = workbookDoc.CreateElement("sheets");
            workBook.AppendChild(sheets);

            // Creating the file for Excel Package
            string targetDir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            FileInfo targetFile = new FileInfo(targetDir + "\\" + FileName + ".xlsx");

            if (targetFile.Exists)
            {
                FileName = FileName + Path.GetRandomFileName();
            }

            ExcelFile = Package.Open(targetDir + "\\" + FileName + ".xlsx", FileMode.Create, FileAccess.ReadWrite);
        }

        private void InitializeSharedStringDoc()
        {
            ssCount = 0;
            sharedStringsDoc = new XmlDocument();

            //Get a reference to the root node, and then add the XML declaration.
            XmlElement ssRoot = sharedStringsDoc.DocumentElement;
            XmlDeclaration ssxmldecl = sharedStringsDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            sharedStringsDoc.InsertBefore(ssxmldecl, ssRoot);

            //Create and append the sst node.
            sstNode = sharedStringsDoc.CreateElement("sst");
            sstNode.SetAttribute("xmlns", "http://schemas.openxmlformats.org/" + "spreadsheetml/2006/main");
            sharedStringsDoc.AppendChild(sstNode);
        }

        public void AddSheet(string SheetName = "Sheet")
        {
            Sheets.Add(new ExcelSheet(SheetName));

            if (SheetName == "Sheet")
            {
                SheetName += Sheets.Count.ToString();
            }

            //Create and append the <sheet> node to the <sheets> node.
            XmlElement sheet = workbookDoc.CreateElement("sheet");
            sheet.SetAttribute("name", SheetName);
            sheet.SetAttribute("sheetId", Sheets.Count.ToString());
            sheet.SetAttribute("id", "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships", "rId" + Sheets.Count.ToString());
            sheets.AppendChild(sheet);
        }

        /// <summary>
        /// Method creates the shared strings document for the excel file generated
        /// </summary>
        public void SaveChanges()
        {
            
            foreach (ExcelSheet sheet in Sheets)
            {
                int rowCount = 0;

                if (sheet.autoFilter)
                {
                    // Adding element for AutoFilter
                    XmlElement autoFilterElement = sheet.workSheetDoc.CreateElement("autoFilter");
                    string startRange = sheet.Rows.First().Columns.First().cellAddress;
                    string endRange = sheet.Rows.Last().Columns.Last().cellAddress;
                    autoFilterElement.SetAttribute("ref", startRange + ":" + endRange);

                    sheet.workSheetDoc.DocumentElement.AppendChild(autoFilterElement);
                }

                foreach (ExcelRow row in sheet.Rows)
                {
                    rowCount++;
                    int columnCount = 0;

                    foreach (ExcelColumn cell in row.Columns)
                    {
                        columnCount++;

                        //Check for availability of value in a cell
                        if (!string.IsNullOrEmpty(cell.Value))
                        {
                            //Create the si node
                            XmlElement siNode = sharedStringsDoc.CreateElement("si");

                            //Create and append the t node.
                            XmlElement tNode = sharedStringsDoc.CreateElement("t");
                            tNode.InnerText = cell.Value;
                            siNode.AppendChild(tNode);

                            //append the si node to sharedStrings Document
                            sharedStringsDoc.ChildNodes[1].AppendChild(siNode);

                            cell.vNode.InnerText = ssCount.ToString();

                            row.RowDoc.FirstChild.ReplaceChild(row.RowDoc.ImportNode(cell.ColDoc.FirstChild, true), row.RowDoc.FirstChild.ChildNodes[columnCount - 1]);
                            sheet.workSheetDoc.ChildNodes[1].FirstChild.ReplaceChild(sheet.workSheetDoc.ImportNode(row.RowDoc.FirstChild, true), sheet.workSheetDoc.ChildNodes[1].FirstChild.ChildNodes[rowCount - 1]);

                            ssCount++;
                        }
                    }
                }
            }

            sstNode.SetAttribute("count", ssCount.ToString());
            sstNode.SetAttribute("uniqueCount", ssCount.ToString());

            XmlDocument tempWorkBook = workbookDoc;
            
            XmlDocument tempSharedStrings = sharedStringsDoc;
            
            XmlDocument tempWorkSheet = Sheets.Last().workSheetDoc;
        }

        public void Download()
        {
            SaveChanges();

            if (ExcelFile != null)
            {
                AddExcelParts();
            }

            if (ExcelFile != null)
            {
                ExcelFile.Flush();
                ExcelFile.Close();
            }
        }

        private void AddExcelParts()
        {
            #region Add excel part of Workbook

            string nsWorkbook = "application/vnd.openxmlformats-" + "officedocument.spreadsheetml.sheet.main+xml";
            string workbookRelationshipType = "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/" + "officeDocument";
            Uri workBookUri = PackUriHelper.CreatePartUri(new Uri("xl/workbook.xml", UriKind.Relative));

            //Create the workbook part.
            PackagePart wbPart = ExcelFile.CreatePart(workBookUri, nsWorkbook);

            //Write the workbook XML to the workbook part.
            Stream workbookStream = wbPart.GetStream(FileMode.Create, FileAccess.Write);
            workbookDoc.Save(workbookStream);

            //Create the relationship for the workbook part.
            ExcelFile.CreateRelationship(workBookUri, TargetMode.Internal, workbookRelationshipType, "rId1");

            #endregion

            #region Add excel part for Worksheet

            foreach (ExcelSheet sheet in Sheets)
            {
                string nsWorksheet = "application/vnd.openxmlformats-" + "officedocument.spreadsheetml.worksheet+xml";
                string worksheetRelationshipType = "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships/worksheet";
                Uri workSheetUri = PackUriHelper.CreatePartUri(new Uri("xl/worksheets/" + sheet.SheetName + ".xml", UriKind.Relative));

                //Create the workbook part.
                PackagePart wsPart = ExcelFile.CreatePart(workSheetUri, nsWorksheet);

                //Write the workbook XML to the workbook part.
                Stream worksheetStream = wsPart.GetStream(FileMode.Create, FileAccess.Write);
                sheet.workSheetDoc.Save(worksheetStream);

                //Create the relationship for the workbook part.
                Uri wsworkbookPartUri = PackUriHelper.CreatePartUri(new Uri("xl/workbook.xml", UriKind.Relative));
                PackagePart wsworkbookPart = ExcelFile.GetPart(wsworkbookPartUri);
                wsworkbookPart.CreateRelationship(workSheetUri, TargetMode.Internal, worksheetRelationshipType, "rId1");
            }

            #endregion

            #region Add excel part for Shared Strings

            string nsSharedStrings = "application/vnd.openxmlformats-officedocument" + ".spreadsheetml.sharedStrings+xml";
            string sharedStringsRelationshipType = "http://schemas.openxmlformats.org" + "/officeDocument/2006/relationships/sharedStrings";
            Uri sharedStringsUri = PackUriHelper.CreatePartUri(new Uri("xl/sharedStrings.xml", UriKind.Relative));

            //Create the workbook part.
            PackagePart sharedStringsPart = ExcelFile.CreatePart(sharedStringsUri, nsSharedStrings);

            //Write the workbook XML to the workbook part.
            Stream sharedStringsStream = sharedStringsPart.GetStream(FileMode.Create, FileAccess.Write);
            sharedStringsDoc.Save(sharedStringsStream);

            //Create the relationship for the workbook part.
            Uri ssworkbookPartUri = PackUriHelper.CreatePartUri(new Uri("xl/workbook.xml", UriKind.Relative));
            PackagePart ssworkbookPart = ExcelFile.GetPart(ssworkbookPartUri);
            ssworkbookPart.CreateRelationship(sharedStringsUri, TargetMode.Internal, sharedStringsRelationshipType, "rId2");

            #endregion
        }

        #endregion
    }
}
