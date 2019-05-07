using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Materials_Processor;

namespace MTO_Report_Processor
{
    public partial class IsoLogForm : Form
    {
        public DataTable Table { get; set; }
        public DataTable dt = new DataTable();
        public DataTable isologtable = new DataTable();
        List<string> TList = new List<string>();
      public  List<string> ISOList = new List<string>();
        string selection;

        PD_EDWDataSetTableAdapters.isoLogTableAdapter isologchecker = new PD_EDWDataSetTableAdapters.isoLogTableAdapter();
        PD_EDWDataSetTableAdapters.jobsTableAdapter jobschcker = new PD_EDWDataSetTableAdapters.jobsTableAdapter();

        public IsoLogForm()
        {
            InitializeComponent();
        }

        private void IsoLogForm_Load(object sender, EventArgs e)
        {
            PD_EDWDataSet.jobsDataTable jDT   = new PD_EDWDataSet.jobsDataTable();
           this.jobschcker.Fill(jDT);
            jobsBindingSource.DataSource = jDT;

        }
        

    private void button1_Click(object sender, EventArgs e)
        {
            var message = "";
            foreach (DataGridViewCell r in dataGridView1.SelectedCells)
            {
                TList.Add(r.Value.ToString());
                
            }
            foreach (String iso in TList)    
            {
                //isologchecker.getisolistby(selection, iso);
                DataTable dt = isologchecker.GetDataByisologlist(selection, iso);
                foreach (DataRow row in dt.Rows)
                {
                    ISOList.Add(row["isoLog_refDwg"].ToString());
                }
                message = string.Join(Environment.NewLine, ISOList);
            }

            Form1 form = new Form1();
            form.passList(this.ISOList);
        }

        private void comboBox1_SelectionChangeCommitted(object sender, EventArgs e)
        {
            string selection = comboBox1.SelectedText.ToString();
            //dataGridView1.Rows.Clear();
            dataGridView1.DataSource = isologchecker.GetData(selection);
            dataGridView1.Refresh();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            selection = comboBox1.SelectedValue.ToString();
            //dataGridView1.Rows.Clear();
            dataGridView1.DataSource = isologchecker.GetData(selection);
            dataGridView1.Refresh();
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
        }
    }
}
