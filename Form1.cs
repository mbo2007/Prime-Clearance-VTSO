using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Tools.Excel;


namespace PrimeAddin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            DataTable dt = new DataTable();
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
