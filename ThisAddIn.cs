using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

namespace PrimeAddin
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
        // edit to suit needs
        // perhaps should return a set of rows/addresses? 
        internal void DemoFind()
        {
            Excel.Worksheet sheet = Application.ActiveSheet;
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            // last is last used range before cells get empty
            Excel.Range last = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            Excel.Range range = sheet.get_Range("A1", last);

            Excel.Range InsuredNames = range.Columns["C"];
             
            // keep these two for later
            int lastUsedRow = last.Row;
            int lastUsedColumn = last.Column;

            // this method searches through InsuredNames (column C) for the first param
            currentFind = InsuredNames.Find("construction", missing,
                Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                missing, missing);

            while (currentFind != null)
            {
                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                    // use EntireRow to return entire row of cell containing "construction"
                    // Cell.EntireRow.Row to get index? need a way to launch datagridview...
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                // action taken after match(es) found
                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);

                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

               // for (int i; i < lastUsedRow; i++)
                //{
                    //Form1.dataGridView1.Rows[i].Cells["Column1"].Value = sheet.Cells[i + 1, 1].Value;
                    //dataGridView1.Rows[i].Cells["Column2"].Value = sheet.Cells[i + 1, 2].Value;
                    //dataGridView1.Rows.Add(sheet.Cells[i + 1, 1].Value, sheet.Cells[i + 1, 2].Value);
                //}

                // FindNext uses previous search settings to repeat search
                currentFind = InsuredNames.FindNext(currentFind);
                // necessary to loop thru FindNext until finished with range?
            }
        }
    }
}
