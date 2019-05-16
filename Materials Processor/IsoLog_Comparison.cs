using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MaterialSkin;
using MaterialSkin.Controls;
using System.Drawing;


namespace MTO_Report_Processor
{
    public partial class IsoLog_Comparison : MaterialForm
    {
        bool switcher = MTO_Report_Processor.Properties.Settings.Default.Theme;
        public IsoLog_Comparison()
        {

            InitializeComponent();
            bool switcher = MTO_Report_Processor.Properties.Settings.Default.Theme;
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

        private void IsoLog_Comparison_Load(object sender, EventArgs e)
        {
            var isolist = MTO_Report_Processor.Properties.Settings.Default.IsoList.Cast<string>().ToList();    //this is the line Jeremy added.


            DataTable testtable = new DataTable();
            testtable.Columns.Add("ISO-Log-RfDwg");

            DataTable testtable2 = new DataTable();
            testtable2.Columns.Add("Files-RfDwg");


            var list1 = new List<string>();
            var list2 = new List<string>();
            var list3 = new List<string>();
            foreach (string blah in MTO_Report_Processor.Properties.Settings.Default.IsoList)
            {
                testtable.Rows.Add(blah);
            }
            foreach (string blah2 in MTO_Report_Processor.Properties.Settings.Default.isotakeofflist)
            {
                testtable2.Rows.Add(blah2);
            }



            dataGridView1.DataSource = testtable;
            dataGridView2.DataSource = testtable2;




            dataGridView1.AutoResizeColumns();
            dataGridView2.AutoResizeColumns();
            dataGridView1.Rows[0].Selected = false;
            dataGridView2.Rows[0].Selected = false;

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

                dataGridView2.DefaultCellStyle.BackColor = Color.DimGray;
                dataGridView2.GridColor = Color.WhiteSmoke;
                dataGridView2.DefaultCellStyle.ForeColor = Color.White;
                dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkGray;
                dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.DarkGray;
                dataGridView2.RowHeadersDefaultCellStyle.ForeColor = Color.White;
                dataGridView2.EnableHeadersVisualStyles = false;

                dataGridView1.DefaultCellStyle.BackColor = Color.Red;
                dataGridView1.DefaultCellStyle.ForeColor = Color.White;
                dataGridView2.DefaultCellStyle.BackColor = Color.Red;
                dataGridView2.DefaultCellStyle.ForeColor = Color.White;


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

                dataGridView2.DefaultCellStyle.BackColor = Color.White;
                dataGridView2.GridColor = Color.Black;
                dataGridView2.DefaultCellStyle.ForeColor = Color.Black;
                dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.White;
                dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
                dataGridView2.RowHeadersDefaultCellStyle.BackColor = Color.White;
                dataGridView2.RowHeadersDefaultCellStyle.ForeColor = Color.Black;
                dataGridView2.EnableHeadersVisualStyles = false;





            }

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void groupBox1_Enter(object sender, EventArgs e)
        {

        }

        private void IsoLog_Comparison_Shown(object sender, EventArgs e)
        {

            foreach (DataGridViewRow row2 in dataGridView2.Rows)
            {
                foreach (DataGridViewRow row1 in dataGridView1.Rows)
                {
                    if (row1.Cells[0].Value.ToString() == row2.Cells[0].Value.ToString())
                    {
                        row2.Cells[0].Style.BackColor = Color.Green;
                        row2.Cells[0].Style.ForeColor = Color.White;
                    }
                }
            }


            foreach (DataGridViewRow row2 in dataGridView1.Rows)
            {
                foreach (DataGridViewRow row1 in dataGridView2.Rows)
                {
                    if (row1.Cells[0].Value.ToString() == row2.Cells[0].Value.ToString())
                    {
                        row2.Cells[0].Style.BackColor = Color.Green;
                        row2.Cells[0].Style.ForeColor = Color.White;
                    }
                }
                dataGridView1.AutoResizeColumns();
                dataGridView2.AutoResizeColumns();
            }



            //foreach (DataGridViewRow row1 in dataGridView1.Rows)
            //{
            //    foreach (DataGridViewRow row2 in dataGridView2.Rows)
            //    {
            //        if (!row2.Cells[0].Value.ToString().Contains(row1.Cells[0].Value.ToString()))
            //            row2.Cells[0].Style.ForeColor = Color.Red;
            //    }
            //}
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
