using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using PrimeAddin;


namespace PrimeAddin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DataTable matchingRows = Globals.ThisAddIn.DemoFind();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }     

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            //need to share things in this class with ThisAddIn            
        }
    }

}
