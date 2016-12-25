using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace OpenExcel
{
    /// <summary>
    /// Class for Excel Column
    /// </summary>
    public class ExcelColumn
    {
        public int ColNum { get; private set; }
        
        internal string ExcelCol { get; set; }
        internal XmlElement cNode { get; private set; }
        internal XmlElement vNode { get; set; }
        internal XmlDocument ColDoc = null;
        
        // Public properties
        
        // Storing the address of the cell
        public string cellAddress { get; set; }
        
        // Storing the Value stored in the cell
        public string Value { get; set; }

        public bool isReadOnly { get; set; }
        
        public bool TakeValueFromSharedString { get; set; }

        private static char[] Alphabets = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();

        #region Constructors

        public ExcelColumn()
        {
            ColDoc = new XmlDocument();
            Value = "";

            cNode = ColDoc.CreateElement("c");
            ColDoc.AppendChild(cNode);

            vNode = ColDoc.CreateElement("v");
            vNode.InnerText = Value.ToString();
            cNode.AppendChild(vNode);
        }

        internal ExcelColumn(int ColNum)
        {
            this.ColNum = ColNum + 1;
            CalculateExcelCol();
            Value = "";
            InitializeColumn();
        }

        #endregion

        #region Methods

        private void InitializeColumn()
        {
            ColDoc = new XmlDocument();

            cNode = ColDoc.CreateElement("c");
            ColDoc.AppendChild(cNode);

            vNode = ColDoc.CreateElement("v");
            cNode.AppendChild(vNode);
        }

        internal void setColumnNumber(int colNum)
        {
            this.ColNum = colNum + 1;
            CalculateExcelCol();
        }

        private void CalculateExcelCol()
        {
            if (string.IsNullOrEmpty(cellAddress))
            {
                int count = 0;
                int num = ColNum;

                while (num > 26)
                {
                    num -= 26;
                    count++;
                }

                if (count > 0)
                {
                    ExcelCol = Alphabets[count - 1].ToString() + Alphabets[num - 1].ToString();
                }
                else
                    ExcelCol = Alphabets[num - 1].ToString();
            }
        }

        #endregion
    }
}
