using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using PrimeAddin;


namespace PrimeAddin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void Launchbtn_Click(object sender, RibbonControlEventArgs e)
        {
            // must find a way to launch datagridview/ share local variables
            Form1 form1 = new Form1();
            //use params of constructor to pass rows into datagridview? cycle/iterate thru arguments
            form1.Show();
        }


    }
}
