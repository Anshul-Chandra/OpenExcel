using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OpenExcel
{
    /// <summary>
    /// Class for Excel Row
    /// </summary>
    public class ExcelRow
    {
        public int RowNum { get; private set; }
        public List<ExcelColumn> Columns { get; set; }
        public XmlElement rNode { get; private set; }
        internal XmlDocument RowDoc = null;

        #region Constructors

        public ExcelRow()
        {
            RowDoc = new XmlDocument();
            Columns = new List<ExcelColumn>();
            rNode = RowDoc.CreateElement("row");
            RowDoc.AppendChild(rNode);
        }

        internal ExcelRow(int RowNum)
        {
            this.RowNum = RowNum + 1;
            Columns = new List<ExcelColumn>();
            InitializeRow();
        }

        #endregion

        #region Methods

        public void InitializeRow()
        {
            RowDoc = new XmlDocument();

            rNode = RowDoc.CreateElement("row");
            rNode.SetAttribute("r", RowNum.ToString());
            rNode.SetAttribute("spans", "1:1");
            RowDoc.AppendChild(rNode);
        }

        internal void AddRowNum(int RowNum)
        {
            this.RowNum = RowNum;
        }

        public void AddColumn()
        {
            Columns.Add(new ExcelColumn(Columns.Count));

            rNode.AppendChild(RowDoc.ImportNode(Columns.Last().cNode, true));
            SetColumnAttribute();
        }

        public void AddColumn(ExcelColumn column)
        {
            column.setColumnNumber(Columns.Count);

            Columns.Add(column);

            rNode.AppendChild(RowDoc.ImportNode(Columns.Last().cNode, true));
            SetColumnAttribute();
        }

        private void SetColumnAttribute()
        {
            Columns.Last().cellAddress = Columns.Last().ExcelCol + RowNum.ToString();
            Columns.Last().cNode.SetAttribute("r", Columns.Last().cellAddress);
            Columns.Last().cNode.SetAttribute("t", "s");
        }

        #endregion
    }
}
