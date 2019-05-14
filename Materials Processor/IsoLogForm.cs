using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using Materials_Processor;

using MaterialSkin;
using MaterialSkin.Controls;
using System.Drawing;

namespace MTO_Report_Processor
{
    public partial class IsoLogForm : MaterialForm
    {

        bool switcher = MTO_Report_Processor.Properties.Settings.Default.Theme;

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

            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;

            // Configure color schema
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.Blue800, Primary.Blue900,
                Primary.Blue500, Accent.LightBlue200,
                TextShade.WHITE

            );
        }

        private void IsoLogForm_Load(object sender, EventArgs e)
        {
            PD_EDWDataSet.jobsDataTable jDT   = new PD_EDWDataSet.jobsDataTable();
           this.jobschcker.Fill(jDT);
            jobsBindingSource.DataSource = jDT;

            string lastfive = MTO_Report_Processor.Properties.Settings.Default.JobNum.Substring(2, 5);

            dataGridView1.DataSource = isologchecker.GetData(lastfive);
            dataGridView1.Refresh();
            dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.RowCount - 1;
            ThemeChanger();

        }



        private void ThemeChanger()
        {
            if (switcher == false)
            {
                MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
                materialSkinManager.AddFormToManage(this);
                materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;

                // Configure color schema
                materialSkinManager.ColorScheme = new ColorScheme(
                    Primary.BlueGrey800, Primary.BlueGrey900,
                    Primary.BlueGrey500, Accent.LightBlue200,
                    TextShade.WHITE
                );
                dataGridView1.DefaultCellStyle.BackColor = Color.DimGray;
                dataGridView1.GridColor = Color.WhiteSmoke;
                dataGridView1.DefaultCellStyle.ForeColor = Color.White;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.DarkGray;
                dataGridView1.RowHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView1.EnableHeadersVisualStyles = false;



            }
            if (switcher == true)
            {
                MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
                materialSkinManager.AddFormToManage(this);
                materialSkinManager.Theme = MaterialSkinManager.Themes.LIGHT;

                // Configure color schema
                materialSkinManager.ColorScheme = new ColorScheme(
                    Primary.Blue800, Primary.Blue900,
                    Primary.Blue500, Accent.LightBlue200,
                    TextShade.WHITE
                );



                dataGridView1.DefaultCellStyle.BackColor = Color.White;
                dataGridView1.GridColor = Color.Black;
                dataGridView1.DefaultCellStyle.ForeColor = Color.Black;
                dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
                dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                dataGridView1.RowHeadersDefaultCellStyle.BackColor = Color.White;
                dataGridView1.RowHeadersDefaultCellStyle.ForeColor = Color.Black;
                dataGridView1.EnableHeadersVisualStyles = false;

            }

        }




        private void button1_Click_1(object sender, EventArgs e)
        {
            MTO_Report_Processor.Properties.Settings.Default.IsoList.Clear();
            string lastfive = MTO_Report_Processor.Properties.Settings.Default.JobNum.Substring(2, 5);
            var message = "";
            foreach (DataGridViewCell r in dataGridView1.SelectedCells)
            {
                TList.Add(r.Value.ToString());

            }
            foreach (String trans in TList)
            {
                //isologchecker.getisolistby(selection, iso);

                DataTable dt = isologchecker.GetDataByisologlist(lastfive, trans);
                foreach (DataRow row in dt.Rows)
                {
                    MTO_Report_Processor.Properties.Settings.Default.IsoList.Add(row["isoLog_refDwg"].ToString());
                }
                //message = string.Join(Environment.NewLine, MTO_Report_Processor.Properties.Settings.Default.IsoList);
                //MessageBox.Show(message.ToString());
            }

            //Form1 form = new Form1();
            //form.passList(this.ISOList);

            IsoLog_Comparison form2 = new IsoLog_Comparison();
            form2.Show();

            this.Close();
        }
    }
}
