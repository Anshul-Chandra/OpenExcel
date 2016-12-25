using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OpenExcel
{
    /// <summary>
    /// Class for Excel Sheet
    /// </summary>
    public class ExcelSheet
    {
        public string SheetName { get; set; }
        public List<ExcelRow> Rows { get; set; }
        public XmlDocument workSheetDoc { get; set; }
        public XmlElement sheetData { get; set; }

        public bool autoFilter { internal get; set; }

        #region Constructors

        public ExcelSheet()
        {
            this.SheetName = "Sheet";
            Rows = new List<ExcelRow>();
            InitializeWorkSheetDocument();
        }

        public ExcelSheet(string SheetName)
        {
            this.SheetName = SheetName;
            Rows = new List<ExcelRow>();
            InitializeWorkSheetDocument();
        }

        #endregion

        #region Methods

        private void InitializeWorkSheetDocument()
        {
            //Create a new XML document for the worksheet.
            workSheetDoc = new XmlDocument();

            //Get a reference to the root node, and then add
            //the XML declaration.
            XmlElement wsRoot = workSheetDoc.DocumentElement;
            XmlDeclaration wsxmldecl = workSheetDoc.CreateXmlDeclaration("1.0", "UTF-8", "yes");
            workSheetDoc.InsertBefore(wsxmldecl, wsRoot);


            //Create and append the worksheet node to the document.
            XmlElement workSheet = workSheetDoc.CreateElement("worksheet");
            workSheet.SetAttribute("xmlns", "http://schemas.openxmlformats.org/" + "spreadsheetml/2006/main");
            workSheet.SetAttribute("xmlns:r", "http://schemas.openxmlformats.org/" + "officeDocument/2006/relationships");
            workSheetDoc.AppendChild(workSheet);

            //Create and add the sheetData node.
            sheetData = workSheetDoc.CreateElement("sheetData");
            workSheet.AppendChild(sheetData);
        }

        public void AddRow()
        {
            Rows.Add(new ExcelRow(Rows.Count));

            sheetData.AppendChild(workSheetDoc.ImportNode(Rows.Last().rNode, true));
        }

        public void AddRow(ExcelRow row)
        {
            row.AddRowNum(Rows.Count);
            Rows.Add(row);

            sheetData.AppendChild(workSheetDoc.ImportNode(Rows.Last().rNode, true));

            AddAttributesToRow();
        }

        private void AddAttributesToRow()
        {
            Rows.Last().rNode.SetAttribute("r", Rows.Last().RowNum.ToString());
            Rows.Last().rNode.SetAttribute("spans", "1:1");
        }

        #endregion
    }
}
