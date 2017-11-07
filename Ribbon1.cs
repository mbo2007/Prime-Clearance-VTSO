using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;


namespace PrimeAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Launchbtn_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.DemoFind();
            // must find a way to launch datagridview/ share local variables
            Form1 form1 = new Form1();
            form1.Show();
        }


    }
}
