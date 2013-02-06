using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration.CustomUI;
using ExcelDna.Integration;
using System.Windows.Forms;

namespace DDSC
{
    public class Main : IExcelAddIn
    {

        public void AutoClose()
        {
            //MessageBox.Show("Closing!");
        }

        public void AutoOpen()
        {
            //MessageBox.Show("Starting");
        }

        [ExcelCommand(MenuName="DDSC", MenuText = "Generate UUID")]
        public static void GenerateGuid()
        {
            ExcelReference selectedCells = (ExcelReference)XlCall.Excel(XlCall.xlfSelection);
            for (int row = selectedCells.RowFirst; row <= selectedCells.RowLast; ++row)
            {
                for (int column = selectedCells.ColumnFirst; column <= selectedCells.ColumnLast; ++column)
                {
                    string guid = Guid.NewGuid().ToString().ToUpper();
                    ExcelReference activeCell = new ExcelReference(row, column);
                    activeCell.SetValue(guid);
                }
            }
        }
    }
}
