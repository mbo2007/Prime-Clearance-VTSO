using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using PrimeAddin;

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
         internal System.Data.DataTable DemoFind()
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            //this seems messy/unwieldy to add each column like this... mark for future consideration
            dt.Columns.Add("Date Submitted");
            dt.Columns.Add("Effective Date");
            dt.Columns.Add("Insured Name");
            dt.Columns.Add("U/W");
            dt.Columns.Add("A/U");
            dt.Columns.Add("Broker");
            dt.Columns.Add("Broker Name");
            dt.Columns.Add("Retail Broker");
            dt.Columns.Add("Practice Policy");
            dt.Columns.Add("Designated Project");
            dt.Columns.Add("Project Address");
            dt.Columns.Add("GC");
            dt.Columns.Add("Owner's Interest");
            dt.Columns.Add("Owner/GC");
            dt.Columns.Add("OCP");
            dt.Columns.Add("Trade");
            dt.Columns.Add("Products");
            dt.Columns.Add("Other");
            dt.Columns.Add("Declined/Blocked/DEAD");
            dt.Columns.Add("Indication");
            dt.Columns.Add("SNIC USIC CBIC CBSIC");
            dt.Columns.Add("PRIMARY Quoted (Yes)");
            dt.Columns.Add("EXCESS Quoted (Yes)");
            dt.Columns.Add("Date Quote Sent");
            dt.Columns.Add("BOUND");
            dt.Columns.Add("BOUND Policy Number");
            dt.Columns.Add("BOUND Policy Premium");
            dt.Columns.Add("TRIA");
            dt.Columns.Add("XPCO");
            dt.Columns.Add("OL&T");
            dt.Columns.Add("Other AP's");
            dt.Columns.Add("TOTAL PREMIUM CHARGED");
            dt.Columns.Add("In-House Loss Fee");
            dt.Columns.Add("BOUND Policy Sent to Broker");
            dt.Columns.Add("NOTES");

            Excel.Worksheet sheet = Globals.ThisAddIn.Application.ActiveSheet;
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;

            List<Excel.Range> matchingRows = new List<Excel.Range>();

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
                matchingRows.Add(currentFind.EntireRow);

                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

                // for (int i; i < lastUsedRow; i++)
                //{
                //Form1.dataGridView1.Rows[i].Cells["Column1"].Value = sheet.Cells[i + 1, 1].Value;
                //dataGridView1.Rows[i].Cells["Column2"].Value = sheet.Cells[i + 1, 2].Value;
                //dataGridView1.Rows.Add(sheet.Cells[i + 1, 1].Value, sheet.Cells[i + 1, 2].Value);
                //}

                // at end, before executing FindNext, CREATE ARRAY CONTAINING ALL VALUES OF ROW
                // PASS THIS ARRAY INTO FORM BY CALLING METHOD IN FORM THAT ACCEPTS IN PARAMS/ CREATES NEW ROW IN DT WHEN CALLED

                // FindNext uses previous search settings to repeat search
                currentFind = InsuredNames.FindNext(currentFind);
                // necessary to loop thru FindNext until finished with range?
            }

            //could be simplified/consolidated into while(currentFind) loop - could remove unnecessary matchingRows array?
            for (int r = 0; r < matchingRows.Count; r++)
            {
                System.Data.DataRow dtRow = dt.NewRow();
                for (int c = 1; c < 35; c++)
                {
                    dtRow[c] = sheet.Cells[matchingRows[r].Row, c].Value;
                }
                dt.Rows.Add(dtRow);
            }
            return dt;
        } 
        
    }
}
