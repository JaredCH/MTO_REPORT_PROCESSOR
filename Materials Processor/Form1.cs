using System;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Windows.Forms;
using System.Diagnostics;
using System.Reflection;
using System.Net;
using System.Net.Mail;
using System.Runtime.InteropServices;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Configuration;
using System.Linq;
using System.Text;
using Materials_Processor;
using System.Collections.Generic;

using Excel = Microsoft.Office.Interop.Excel;
using MTO_Report_Processor;
using Microsoft.Office.Interop.Access;
using Microsoft;
using ExcelDataReader;
using ExcelDataReader.Core;
using ExcelDataReader.Exceptions;
using ExcelDataReader.Log;
using MaterialSkin;
using MaterialSkin.Controls;


namespace Materials_Processor
{
        public partial class Form1 : MaterialForm
    {
        //PD_EDWDataSet.jobsTableAdapter jobschcker = new PD_EDWDataSet.jobsTableAdapter();
        MTO_Report_Processor.PD_EDWDataSet1TableAdapters.isoLogTableAdapter isologchecker = new MTO_Report_Processor.PD_EDWDataSet1TableAdapters.isoLogTableAdapter();





        List<string> new_ISOS = new List<string>();
        public Form1(List<string>ISOS)
        {
            new_ISOS = ISOS;
        }

        string jobnum, trans;

        DataTable STOTABLE = new DataTable();
        string logPath = "";
        string stopath = "";



        public Form1()
        {



            
            InitializeComponent();
            MaterialForm f = new MaterialForm();
            MaterialSkinManager materialSkinManager = MaterialSkinManager.Instance;
            materialSkinManager.AddFormToManage(this);
            materialSkinManager.Theme = MaterialSkinManager.Themes.DARK;

            // Configure color schema
            materialSkinManager.ColorScheme = new ColorScheme(
                Primary.BlueGrey800, Primary.BlueGrey900,
                Primary.BlueGrey500, Accent.LightBlue200,
                TextShade.WHITE
            );


            typeof(DataGridView).InvokeMember("DoubleBuffered",
            BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
            null, this.dataGridView1, new object[] { true });
            typeof(DataGridView).InvokeMember("DoubleBuffered",
BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
null, this.dataGridView2, new object[] { true });
            System.IO.Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MTO Report Processor");
        }


        string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MTO Report Processor";
        static DataTable GetDataTableFromCsv(string path, bool isFirstRowHeader)

        {
            string header = isFirstRowHeader ? "Yes" : "No";

            string pathOnly = Path.GetDirectoryName(path);
            string fileName = Path.GetFileName(path);

            string sql = @"SELECT * FROM [" + fileName + "]";

            using (OleDbConnection connection = new OleDbConnection(
                      @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathOnly +
                      ";Extended Properties=\"Text;HDR=" + header + ";FMT=Delimited($)\""))
            using (OleDbCommand command = new OleDbCommand(sql, connection))
            using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
            {
                DataTable dataTable = new DataTable();
                dataTable.Locale = CultureInfo.CurrentCulture;
                adapter.Fill(dataTable);
                return dataTable;
            }
        }


        public DataTable ConvertToDataTable(string filePath, int numberOfColumns)
        {
            DataTable tbl = new DataTable();
            DataTable tbl2 = new DataTable();
            DataTable Final = new DataTable();
            tbl.Columns.Clear();
            tbl2.Columns.Clear();
            Final.Columns.Clear();

            tbl.Columns.Add("Test");
            tbl2.Columns.Add("ISO");
            tbl2.Columns.Add("PCMK");
            tbl2.Columns.Add("PIPING_SPEC");
            tbl2.Columns.Add("UMI");
            tbl2.Columns.Add("SIZE");
            tbl2.Columns.Add("DESCRIPTION");
            tbl2.Columns.Add("ITEM_CODE");
            tbl2.Columns.Add("QTY");
            tbl2.Columns.Add("UnitOFMeasure");
            tbl2.Columns.Add("GROUP");
            tbl2.Columns.Add("SOURCE");
            tbl2.Columns.Add("SIZE1");
            tbl2.Columns.Add("SIZE2");
            tbl2.Columns.Add("SIZE3");
            tbl2.Columns.Add("CATEGORY");


            Final.Columns.Add("Production_No");
            Final.Columns.Add("Source");
            Final.Columns.Add("Pipeline_Reference");
            Final.Columns.Add("Material Code");
            Final.Columns.Add("Spool Number");
            Final.Columns.Add("Piecemark");
            Final.Columns.Add("Piping_Spec");
            Final.Columns.Add("Item_Code");
            Final.Columns.Add("Size");
            Final.Columns.Add("Description");
            Final.Columns.Add("End_Conditions");
            Final.Columns.Add("Tag");
            Final.Columns.Add("Group");
            Final.Columns.Add("Qty");
            Final.Columns.Add("Qty2");
            Final.Columns.Add("UnitOfMeasure");
            Final.Columns.Add("Long_ID");
            Final.Columns.Add("JDE_Desc");
            Final.Columns.Add("Record_Type");
            Final.Columns.Add("Date");
            Final.Columns.Add("recdate");
            Final.Columns.Add("linenum");
            Final.Columns.Add("revnum");
            Final.Columns.Add("linesize");
            Final.Columns.Add("Index", Type.GetType("System.Double"));

            








            string[] lines = System.IO.File.ReadAllLines(filePath);
            foreach (string line in lines)
            {
                tbl.Rows.Add(line);
            }
            int j = 0;
            foreach (DataRow row in tbl.Rows)
            {
                if (tbl.Rows.IndexOf(row) != 0)
                {
                    string[] data = row["Test"].ToString().Split('$');
                    tbl2.Rows.Add(data[0], data[1], data[2], data[3], data[4], data[5], data[6], data[7], data[8], data[9], data[10], data[11], data[12], data[13], data[14]);
                    j++;
                }
            }
            int indexnum = 0;
            foreach (DataRow row in tbl2.Rows)
            {
                if (row["SIZE1"].Equals("3/4"))
                {
                    row["SIZE1"] = "0.75";
                }
                if (row["SIZE2"].Equals("3/4"))
                {
                    row["SIZE2"] = "0.75";
                }
                if (row["SIZE1"].Equals("1/2"))
                {
                    row["SIZE1"] = "0.5";
                }
                if (row["SIZE2"].Equals("1/2"))
                {
                    row["SIZE2"] = "0.5";
                }
                if (row["SIZE1"].Equals("1/4"))
                {
                    row["SIZE1"] = "0.25";
                }
                if (row["SIZE2"].Equals("1/4"))
                {
                    row["SIZE2"] = "0.25";
                }
                if (row["SIZE1"].Equals("1.1/2"))
                {
                    row["SIZE1"] = "1.5";
                }
                if (row["SIZE2"].Equals("1.1/2"))
                {
                    row["SIZE2"] = "1.5";
                }
                if (row["SIZE1"].Equals("2.1/2"))
                {
                    row["SIZE1"] = "2.5";
                }
                if (row["SIZE2"].Equals("2.1/2"))
                {
                    row["SIZE2"] = "2.5";
                }
                if (row["SIZE1"].Equals("3.1/2"))
                {
                    row["SIZE1"] = "3.5";
                }
                if (row["SIZE2"].Equals("3.1/2"))
                {
                    row["SIZE2"] = "3.5";
                }
                if (row["SIZE1"].Equals("4.1/2"))
                {
                    row["SIZE1"] = "4.5";
                }
                if (row["SIZE2"].Equals("4.1/2"))
                {
                    row["SIZE2"] = "4.5";
                }

                row["SIZE"] = row["SIZE1"] + "x" + row["SIZE2"] + "x0";

                string qty = row["QTY"].ToString();
                if (qty.Contains("'"))
                {
                    try
                    {
                        string[] qties = row["QTY"].ToString().Split('\'');
                        Decimal foot = Convert.ToDecimal(qties[0]);
                        string inch = qties[1].Replace("\"", "");
                        Decimal inchdecimal = Convert.ToDecimal(inch);
                        //MessageBox.Show(inch + "and then" + inchdecimal + "and then" + (inchdecimal / 12));
                        row["QTY"] = decimal.Round((foot + inchdecimal / 12), 2);
                        row["UnitOfMeasure"] = "FT";
                    }
                    catch { }
                }
                double qty1 = 0;
                decimal QTY2S =0;
                try
                {
                    if (row["GROUP"].Equals("PIPE"))
                    {
                        row["UnitOfMeasure"] = "FT";
                        qty1 = Convert.ToDouble(row["QTY"]);
                        QTY2S = decimal.Round(Convert.ToDecimal(qty1 * 1.05), 2);
                    }
                    if (!row["GROUP"].Equals("PIPE"))
                    {
                        row["UnitOfMeasure"] = "EA";
                        qty1 = Convert.ToDouble(row["QTY"]);
                        QTY2S = Convert.ToDecimal(qty1);
                    }
                }
                catch { }


                row["GROUP"] = row["CATEGORY"].ToString() + "_" + row["GROUP"].ToString();

                
                Final.Rows.Add(jobnum, trans, row["ISO"].ToString(),"","", row["PCMK"].ToString(), row["PIPING_SPEC"].ToString(), row["Item_Code"].ToString(), row["SIZE"].ToString(), row["DESCRIPTION"].ToString(), "", "", row["GROUP"].ToString(), row["QTY"].ToString(), QTY2S, row["UnitOfMEasure"].ToString(), "", "", "MI", DateTime.Now.ToString("MM/dd/yyyy"),"","","","", indexnum.ToString());
                indexnum++;
            }
            return Final;
        }





        private void Form1_Load(object sender, EventArgs e)
        {          
            // right click option to set "take off method" options for IDF, PCF, Both, or Manual.
        }



        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void quitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void findAndReplaceRefDwgToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string oldtext = Microsoft.VisualBasic.Interaction.InputBox("Text to Replace", "Find and Replace - RefDwg", "Default");
            string newtext = Microsoft.VisualBasic.Interaction.InputBox("Replacing text with", "Find and Replace - RefDwg", "Default");
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Pipeline_Reference"].Value.ToString();
                    row.Cells["Pipeline_Reference"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }


        private void appendSpecInfoSpecToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string Spec = Microsoft.VisualBasic.Interaction.InputBox("Spec to address", "Append Spec info - Spec", "Default");
            string appendtext = Microsoft.VisualBasic.Interaction.InputBox("What to append", "Append Spec info - Spec", "Default");
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string whatspecweon = row.Cells["Piping_Spec"].Value.ToString();
                    if (whatspecweon == Spec.ToUpper())
                    {
                        row.Cells["DESCRIPTION"].Value = row.Cells["DESCRIPTION"].Value.ToString() + appendtext.ToUpper();
                    }
                }
                catch
                { }

            }
        }

        private void isoLogCheckToolStripMenuItem_Click(object sender, EventArgs e)
        {
            label3.Visible = true;
            backgroundWorker1.RunWorkerAsync();
        }

        private void saveToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DataTable dtFromGrid = new DataTable();
            dtFromGrid = dataGridView1.DataSource as DataTable;
            DataSet ds = new DataSet();
            ds.Tables.Add(dtFromGrid);
            ExportDataSetToExcel(ds, jobnum + "_" + trans);
           // MessageBox.Show("Report Saved , you can find the file " + @"V:\MTO\Spoolgen\Reports\Processed_Reports\" +jobnum + "_" + trans + ".xls");

        }



        private void ExportDataSetToExcel(DataSet ds, String template)
        {

            try
            {
                //Creae an Excel application instance
                Excel.Application excelApp = new Excel.Application();
                //Create an Excel workbook instance and open it from the predefined location
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet excelWorkSheet = excelWorkBook.Worksheets[1];
                excelWorkSheet.Name = "DeleteMe";

                foreach (DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    excelWorkSheet = excelWorkBook.Sheets.Add();

                    excelWorkSheet.Name = table.TableName;
                    // Column Headers
                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;

                    }

                    string[,] data = new string[table.Rows.Count, table.Columns.Count];
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            data[j, k] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }

                    excelWorkSheet.Range[excelWorkSheet.Cells[2, 1], excelWorkSheet.Cells[table.Rows.Count + 1, table.Columns.Count]].Value = data;
                }

                excelWorkSheet = excelWorkBook.Worksheets["DeleteMe"];
                excelWorkSheet.Delete();
                excelWorkBook.SaveAs(path + "\\" + template + ".xlsx");
                excelWorkBook.Close();
                excelApp.Quit();
                stopath = path + "\\" + template + ".xlsx";


                DialogResult dresult = new DialogResult();
                dresult = MessageBox.Show("Open Export File?", "Export Created", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dresult == DialogResult.Yes)
                {
                    Process process = new Process();
                    process.StartInfo.FileName = path + "\\" + template + ".xlsx";
                    process.Start();
                }
                else
                {
                    //File.Copy(path + jobnum + "_" + trans, @"V:\MTO\Spoolgen\Reports\Processed_Reports\" + jobnum + "_" + trans + "_Nextgen.xlsx");
                    //MessageBox.Show("Export created:  " + path + "\\" + template + ".xlsx", "Export Created", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        private void ExportDataSetToExcelAndMove(DataSet ds, String template)
        {

            try
            {
                //Creae an Excel application instance
                Excel.Application excelApp = new Excel.Application();
                //Create an Excel workbook instance and open it from the predefined location
                Excel.Workbook excelWorkBook = excelApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);

                Excel.Worksheet excelWorkSheet = excelWorkBook.Worksheets[1];
                excelWorkSheet.Name = "DeleteMe";

                foreach (DataTable table in ds.Tables)
                {
                    //Add a new worksheet to workbook with the Datatable name
                    excelWorkSheet = excelWorkBook.Sheets.Add();

                    excelWorkSheet.Name = table.TableName;
                    // Column Headers
                    for (int i = 1; i < table.Columns.Count + 1; i++)
                    {
                        excelWorkSheet.Cells[1, i] = table.Columns[i - 1].ColumnName;

                    }

                    string[,] data = new string[table.Rows.Count, table.Columns.Count];
                    for (int j = 0; j < table.Rows.Count; j++)
                    {
                        for (int k = 0; k < table.Columns.Count; k++)
                        {
                            data[j, k] = table.Rows[j].ItemArray[k].ToString();
                        }
                    }

                    excelWorkSheet.Range[excelWorkSheet.Cells[2, 1], excelWorkSheet.Cells[table.Rows.Count + 1, table.Columns.Count]].Value = data;
                }

                excelWorkSheet = excelWorkBook.Worksheets["DeleteMe"];
                excelWorkSheet.Delete();
                excelWorkBook.SaveAs(path + "\\" + template + "Nextgen_.xlsx");
                excelWorkBook.Close();
                excelApp.Quit();

                DialogResult dresult = new DialogResult();
                dresult = MessageBox.Show("Open Export File?", "Export Created", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

                if (dresult == DialogResult.Yes)
                {
                    Process process = new Process();
                    MessageBox.Show(path + "\\" + template);
                    process.StartInfo.FileName = path + "\\" + template + "Nextgen_.xlsx";
                    process.Start();
                }
                else
                {
                    //NEW CODE
                    File.Copy(path + "\\" + template + "Nextgen_.xlsx", @"V:\MTO\Spoolgen\Reports\Processed_Reports\" + jobnum + "_" + trans);
                    //   File.Copy(path + jobnum + "_" + trans, @"V:\MTO\Spoolgen\Reports\Processed_Reports\" + jobnum + "_" + trans );
                    MessageBox.Show("Export created:  " + path + "\\" + template + "_Nextgen.xlsx", "Export Created", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message, "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

        }



        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }



        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {

        }


       static int n = 10;
        int j = 0;
        DataTable[] backuptbl = new DataTable[n];
        DataTable dt11 = new DataTable();
        DataColumn column;
        DataRow row;
        public void CreateRecoveryPoint()
        {
            // dataGridView1.Refresh();
            // dataGridView1.DataSource = dataGridView1;
            // dtFromGrid[j] = dataGridView1.DataSource as DataTable;
            // dtFromGrid[j].AcceptChanges();


            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                column = new DataColumn();
                column.ColumnName = col.Name;
                dt11.Columns.Add(column);
            }

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataRow dRow = dt11.NewRow();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    dRow[cell.ColumnIndex] = cell.Value;
                }
                dt11.Rows.Add(dRow);
            }
            undoToolStripMenuItem.DropDownItems.Add("Recovery Point " + j );
            dataGridView1.Refresh();
            backuptbl[j] = dt11;
            j++;
            dt11.Rows.Clear();
            dt11.Columns.Clear();
        }

        private void dataGridViewDSourceChanged(object sender, EventArgs e)
        {
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            CreateRecoveryPoint();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

                }

        private void Undodropdownclick(object sender, ToolStripItemClickedEventArgs e)
        {


        }

        private void removeHighlightedLinesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Int32 selectedRowCount = dataGridView1.Rows.GetRowCount(DataGridViewElementStates.Selected);
            if (selectedRowCount > 0)
            {
                for (int i = 0; i < selectedRowCount; i++)
                {
                    dataGridView1.Rows.RemoveAt(dataGridView1.SelectedRows[0].Index);
                }

            }
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {

        }

        private void fileToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            string vlvfilter = Microsoft.VisualBasic.Interaction.InputBox("What to filter for", "Valve Tags", "Default");
            string tagornum = Microsoft.VisualBasic.Interaction.InputBox("Either type 'Tag' or a count of characters to use from the end of the description", "Valve Tags", "Default");
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    vlvfilter = vlvfilter.ToUpper();
                    tagornum = tagornum.ToUpper();
                    string descfiltercheck = row.Cells["Description"].Value.ToString();
                    if (descfiltercheck.Contains(vlvfilter))
                    {


                        switch (tagornum)
                        {
                            case "TAG":
                                {
                                    row.Cells["Tag"].Value = row.Cells["Item_Code"].Value.ToString();
                                    break;
                                }
                            default:
                                {
                                    int theydidanum = Convert.ToInt32(tagornum);
                                    string itemtagnew = row.Cells["Description"].Value.ToString();
                                    row.Cells["Tag"].Value = itemtagnew.Substring(itemtagnew.Length - theydidanum);
                                    break;
                                }
                        }
                    }
                }


                catch
                { }

            }
        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {


            DataTable dtFromGrid = new DataTable();
            dtFromGrid = (dataGridView1.DataSource as DataTable).Copy();

            List<DataRow> RowsToDelete = new List<DataRow>();
            foreach (DataRow row in dtFromGrid.Rows)
                if (row["Group"].ToString() != null &&
                     row["Group"].ToString().Contains("EREC") || row["Group"].ToString().Contains("SUPPORTS")) RowsToDelete.Add(row);
            foreach (DataRow row in RowsToDelete) dtFromGrid.Rows.Remove(row);
            RowsToDelete.Clear();

            try
            {
                dtFromGrid.Columns.Remove("Material Code");
                dtFromGrid.Columns.Remove("Spool Number");
                dtFromGrid.Columns.Remove("recdate");
                dtFromGrid.Columns.Remove("linenum");
                dtFromGrid.Columns.Remove("revnum");
                dtFromGrid.Columns.Remove("linesize");
                dtFromGrid.Columns.Remove("Index");
                DataSet ds = new DataSet();
                ds.Tables.Clear();
                ds.Tables.Add(dtFromGrid);
                ExportDataSetToExcelAndMove(ds, jobnum + "_" + trans);

            }
            catch { }
        }

        private void sendSTOToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        
        List<string> isoLogList = new List<string>();
        List<string> isoList = new List<string>();
        IEnumerable<string> MissingIDF = new List<string>();
        IEnumerable<string> MissingISOLOG = new List<string>();

        public void passList(List<string> myList)
        {
             isoLogList = myList;
        }


        private void button2_Click_1(object sender, EventArgs e)
        {



        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            string oldtext = ",";
            string newtext = "";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Description"].Value.ToString();
                    row.Cells["Description"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void toolStripMenuItem8_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    row.Cells["Piecemark"].Value = "MK-" + row.Cells["Piecemark"].Value.ToString();
                }
                catch
                { }
            }
        }

        private void generateSTOFromMTOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("ID");
            dt.Columns.Add( "Job_No");
            dt.Columns.Add("Epic_Trans_No");
            dt.Columns.Add("Iso_Rec_Date");
            dt.Columns.Add("Spool_No");
            dt.Columns.Add("Iso_Num");
            dt.Columns.Add("Iso_Rev");
            dt.Columns.Add("Pc_Mk");
            dt.Columns.Add("Client_Desc");
            dt.Columns.Add("Client_Item_Code");
            dt.Columns.Add("Supt_Type");
            dt.Columns.Add("Pipe_Spec");
            dt.Columns.Add("Mtrl_Code");
            dt.Columns.Add("Header_Size");
            dt.Columns.Add("Header_Comp");
            dt.Columns.Add("Combo_Supt");
            dt.Columns.Add("Scope");
            dt.Columns.Add("EPIC_Template");
            dt.Columns.Add("EPIC_Tag#");
            dt.Columns.Add("EPIC_Long_ID");
            dt.Columns.Add("Exist_Detail");
            dt.Columns.Add("Supt_Rev#");
            dt.Columns.Add("Qty_Req");
            dt.Columns.Add("Take_Off_Method");
            dt.Columns.Add("Status");
            dt.Columns.Add("ETD_Trans#");
            dt.Columns.Add("ETD_Trans_Date");
            dt.Columns.Add("DTE_Review_Trans#");
            dt.Columns.Add("DTE_Review_Trans_Date");
            dt.Columns.Add("ETD_Correction_Trans#");
            dt.Columns.Add("ETD_Correction_Trans_Date");
            dt.Columns.Add("DTE_Final_Trans#");
            dt.Columns.Add("DTE_Final_Trans_Date");
            dt.Columns.Add("Date_Iss_for_Fab");
            dt.Columns.Add("Requisition#");
            dt.Columns.Add("Hold");
            dt.Columns.Add("RFI#");
            dt.Columns.Add("Comments");
            dt.Columns.Add("Weight_LBS");
            dt.Columns.Add("Surf_Area_SF");
            dt.Columns.Add("Days_Aged");
            dt.Columns.Add("Item Type");
            dt.Columns.Add("Path");




            

           // dt.Columns.Add("Production_No");
           // dt.Columns.Add("Source");
           // dt.Columns.Add("Pipeline_Reference");
            //dt.Columns.Add("Material Code");
           // dt.Columns.Add("Spool Number");
            //dt.Columns.Add("Piecemark");
            //dt.Columns.Add("Piping_Spec");
           //dt.Columns.Add("Item_Code");
            //dt.Columns.Add("Size");
            //dt.Columns.Add("Description");
            //dt.Columns.Add("End_Conditions");
            //dt.Columns.Add("Tag");
            //dt.Columns.Add("Group");
            //dt.Columns.Add("Qty");
            //dt.Columns.Add("Qty2");
            //dt.Columns.Add("UnitOfMeasure");
            //dt.Columns.Add("Long_ID");
            //dt.Columns.Add("JDE_Desc");
            //dt.Columns.Add("Record_Type");
            //dt.Columns.Add("Date");

            //int i = 0;

         foreach (DataGridViewRow row in dataGridView1.Rows)
                if (row.Cells["Group"].Value != null &&
                     row.Cells["Group"].Value.ToString().Contains("_SUPPORTS"))
                {
                    
                    DataRow toInsert = dt.NewRow();
                    toInsert[8] = row.Cells["Description"].Value.ToString(); ;
                    toInsert[7] = row.Cells["Piecemark"].Value.ToString();
                    toInsert[9] = row.Cells["Item_Code"].Value.ToString();
                    toInsert[11] = row.Cells["Piping_Spec"].Value.ToString();
                    toInsert[12] = row.Cells["Material Code"].Value.ToString();
                    toInsert[4] = row.Cells["Spool Number"].Value.ToString();
                    toInsert[1] = row.Cells["Production_No"].Value.ToString();
                    toInsert[2] = row.Cells["Source"].Value.ToString();

                    toInsert[3] = row.Cells["recdate"].Value.ToString();
                    toInsert[5] = row.Cells["Pipeline_Reference"].Value.ToString();
                    toInsert[6] = row.Cells["revnum"].Value.ToString();
                    toInsert[13] = row.Cells["Size"].Value.ToString();
                    toInsert[22] = row.Cells["Qty"].Value.ToString();
                    // dt.Rows.InsertAt(toInsert, 5);  //(row.Cells["Description"].Value.ToString());

                    //dt.Rows[i]["itemcode"] = row.Cells["Item_Code"].Value.ToString();
                    //i++;
                    dt.Rows.Add(toInsert);
                }
            dt.AcceptChanges();
            this.dataGridView2.DataSource = dt;
            dataGridView2.Refresh();
            
        }

        private void emailSTOToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataTable dtFromGrid = new DataTable();
            dtFromGrid = dataGridView2.DataSource as DataTable;
            DataSet ds = new DataSet();
            ds.Tables.Clear();
            ds.Tables.Add(dtFromGrid);
            ExportDataSetToExcel(ds, jobnum + "_" + trans + "_STO");


            //Outlook.MailItem mailItem = (Outlook.MailItem)
            // this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = jobnum + " T" + trans + " STO Report";
            mailItem.To = "kenneth.smith@EpicPiping.com; sunil.gawli@epicpiping.com; Shailesh.Dabhekar@epicpiping.com";
            mailItem.CC = "andre.naquin@EpicPiping.com; kevin.flores@epicpiping.com; Monty.Cornes@EpicPiping.com; travis.stromain@EpicPiping.com";
            //mailItem.Body = "Please see the attached document.";
            mailItem.Attachments.Add(stopath);//logPath is a string holding path to the log.txt file
            //mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Display(mailItem);
            mailItem.HTMLBody = "Please see the attached document." + mailItem.HTMLBody;

        }

        private void exportSTOToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DataTable dtFromGrid = new DataTable();
            dtFromGrid = dataGridView2.DataSource as DataTable;
            DataSet ds = new DataSet();
            ds.Tables.Clear();
            ds.Tables.Add(dtFromGrid);
            ExportDataSetToExcel(ds, jobnum + "_" + trans + "_STO");
        }

        private void toolStripMenuItem9_Click(object sender, EventArgs e)
        {

        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            DataTable dtFromGrid = new DataTable();
            dtFromGrid = dataGridView1.DataSource as DataTable;
            DataSet ds = new DataSet();
            ds.Tables.Clear();
            ds.Tables.Add(dtFromGrid.Copy());
            ExportDataSetToExcel(ds, jobnum + "_" + trans + "Full_Export");
        }

        private void toolStripMenuItem11_Click(object sender, EventArgs e)
        {
            jobnum = Microsoft.VisualBasic.Interaction.InputBox("Job Number", "New Report Info", "Default");
            trans = Microsoft.VisualBasic.Interaction.InputBox("Transmittal", "New Report Info", "Default");


            string filePath ="";
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "txt files (*.txt)|*.txt|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    //Get the path of specified file
                    filePath = openFileDialog.FileName;

                }

                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    IExcelDataReader reader;
                    reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream);
                                    var conf = new ExcelDataSetConfiguration
                                    {
                                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                                        {
                                            UseHeaderRow = true
                                        }
                                    };

                    var dataSets = reader.AsDataSet(conf);
                    var dataTables = dataSets.Tables[0];
                    dataTables.Columns.Add("Material Code").SetOrdinal(3);
                    dataTables.Columns.Add("Spool Number").SetOrdinal(4);
                    dataTables.Columns.Add("recdate");
                    dataTables.Columns.Add("linenum");
                    dataTables.Columns.Add("revnum");
                    dataTables.Columns.Add("linesize");
                    dataTables.Columns.Add("Index", Type.GetType("System.Double"));
                    dataGridView1.DataSource = dataTables;
                }
            }
            
        }

        private void toolStripMenuItem12_Click(object sender, EventArgs e)
        {
            //try
            //{
                string temppcmk = "";
                string tempiso = "";
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells["Piecemark"].Value != "")
                    {
                        temppcmk = row.Cells["Piecemark"].Value.ToString();
                        tempiso = row.Cells["Pipeline_Reference"].Value.ToString();
                    }
                    //string temppcmktwo = row.Cells["Piecemark"].Value.ToString();
                    if (row.Cells["Piecemark"].Value == "" && row.Cells["Pipeline_Reference"].Value.ToString() == tempiso);
                    {
                        row.Cells["Piecemark"].Value = temppcmk;
                    }
                }
            //}
           // catch { }
        }

        private void pcmkToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtext = "";
                    switch (lastsegment)
                        {
                        case "-1":
                            newtext = "A";
                            break;
                        case "-2":
                            newtext = "B";
                            break;
                        case "-3":
                            newtext = "C";
                            break;
                        case "-4":
                            newtext = "D";
                            break;
                        case "-5":
                            newtext = "E";
                            break;
                        case "-6":
                            newtext = "F";
                            break;
                        case "-7":
                            newtext = "G";
                            break;
                        case "-8":
                            newtext = "H";
                            break;
                        case "-9":
                            newtext = "I";
                            break;


                    }
                    row.Cells["Piecemark"].Value = replaceing.Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2).Replace(lastsegment, firstsegment + newtext);
                }
                catch
                { }
            }
        }

        private void returnToOriginalSortOrderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView1.Sort(dataGridView1.Columns["Index"], System.ComponentModel.ListSortDirection.Ascending);
            }
            catch { }
        }

        private void showListOfSpecsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var vv = dataGridView1.Rows.Cast<DataGridViewRow>()
                           .Where(x => !x.IsNewRow)                   // either..
                           .Where(x => x.Cells["Piping_Spec"].Value != null) //..or or both
                           .Select(x => x.Cells["Piping_Spec"].Value.ToString())
                           .Distinct()
                           .ToList();
            var message = string.Join(Environment.NewLine, vv.ToArray());
            MessageBox.Show(message, "List of Spec's");
        }

        private void showListOfTransmittalsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var vv = dataGridView1.Rows.Cast<DataGridViewRow>()
               .Where(x => !x.IsNewRow)                   // either..
               .Where(x => x.Cells["Source"].Value != null) //..or or both
               .Select(x => x.Cells["Source"].Value.ToString())
               .Distinct()
               .ToList();
            vv.Sort();
            var message = string.Join(Environment.NewLine, vv.ToArray());
            MessageBox.Show(message, "List of Transmittal's");
        }

        private void getCountOfDistinctItemsToolStripMenuItem_Click(object sender, EventArgs e)
        {
                var vv = dataGridView1.Rows.Cast<DataGridViewRow>()
               .Where(x => !x.IsNewRow)                   // either..
               .Where(x => x.Cells["Description"].Value != null)
               .Where(x => x.Cells["Group"].Value.ToString().Contains("FAB"))//..or or both
               .Select(x => x.Cells["Size"].Value.ToString() + x.Cells["Description"].Value.ToString())
               .Distinct()
               .ToList();
            var message = vv.Count().ToString();
            MessageBox.Show(message, "Count of distinct items");
        }

        private void toolStripMenuItem13_Click(object sender, EventArgs e)
        {
            string oldtext = Microsoft.VisualBasic.Interaction.InputBox("Text to Replace", "Find and Replace - Pcmk", "Default");
            string newtext = Microsoft.VisualBasic.Interaction.InputBox("Replacing text with", "Find and Replace - Pcmk", "Default");
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    row.Cells["Piecemark"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void getCountOfISOsToolStripMenuItem_Click(object sender, EventArgs e)
        {
                        var vv = dataGridView1.Rows.Cast<DataGridViewRow>()
            .Where(x => !x.IsNewRow)                   // either..
            .Where(x => x.Cells["Description"].Value != null)
            .Select(x => x.Cells["Pipeline_Reference"].Value.ToString())
            .Distinct()
            .ToList();
            var message = vv.Count().ToString();
            MessageBox.Show(message, "Count of ISO's");
        }

        private void toolStripMenuItem14_Click(object sender, EventArgs e)
        {
            string oldtext = "\"";
            string newtext = "";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Description"].Value.ToString();
                    row.Cells["Description"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void pcmkToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            string oldtext = "-";
            string newtext = " SH.";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Pipeline_Reference"].Value.ToString();
                    row.Cells["Pipeline_Reference"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void toolStripMenuItem6_Click(object sender, EventArgs e)
        {
            string oldtext = "--";
            string newtext = "-";
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    row.Cells["Piecemark"].Value = replaceing.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void pcmkToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtext = "";
                    switch (lastsegment)
                    {
                        case "-1":
                            newtext = "-A";
                            break;
                        case "-2":
                            newtext = "-B";
                            break;
                        case "-3":
                            newtext = "-C";
                            break;
                        case "-4":
                            newtext = "-D";
                            break;
                        case "-5":
                            newtext = "-E";
                            break;
                        case "-6":
                            newtext = "-F";
                            break;
                        case "-7":
                            newtext = "-G";
                            break;
                        case "-8":
                            newtext = "-H";
                            break;
                        case "-9":
                            newtext = "-I";
                            break;


                    }
                    row.Cells["Piecemark"].Value = replaceing.Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2).Replace(lastsegment, firstsegment + newtext);
                }
                catch
                { }
            }
        }

        private void toolStripMenuItem21_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtext = "";
                    switch (lastsegment)
                    {
                        case "-1":
                            newtext = "A";
                            break;
                        case "-2":
                            newtext = "B";
                            break;
                        case "-3":
                            newtext = "C";
                            break;
                        case "-4":
                            newtext = "D";
                            break;
                        case "-5":
                            newtext = "E";
                            break;
                        case "-6":
                            newtext = "F";
                            break;
                        case "-7":
                            newtext = "G";
                            break;
                        case "-8":
                            newtext = "H";
                            break;
                        case "-9":
                            newtext = "I";
                            break;


                    }
                    row.Cells["Piecemark"].Value = replaceing.Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2).Replace(lastsegment, firstsegment + newtext);
                }
                catch
                { }
            }
        }

        private void pcmkToolStripMenuItem3_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtext = "";
                    switch (lastsegment)
                    {
                        case "-1":
                            newtext = "A";
                            break;
                        case "-2":
                            newtext = "B";
                            break;
                        case "-3":
                            newtext = "C";
                            break;
                        case "-4":
                            newtext = "D";
                            break;
                        case "-5":
                            newtext = "E";
                            break;
                        case "-6":
                            newtext = "F";
                            break;
                        case "-7":
                            newtext = "G";
                            break;
                        case "-8":
                            newtext = "H";
                            break;
                        case "-9":
                            newtext = "I";
                            break;


                    }
                    row.Cells["Piecemark"].Value = replaceing.Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2).Replace(lastsegment, firstsegment + newtext);
                }
                catch
                { }
            }
        }

        private void iDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                try
                {
                    row.Cells["Take_Off_Method"].Value = "IDF" ;
                }
                catch
                { }
            }
        }

        private void pCFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                try
                {
                    row.Cells["Take_Off_Method"].Value = "PCF";
                }
                catch
                { }
            }
        }

        private void iDFPCFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                try
                {
                    row.Cells["Take_Off_Method"].Value = "IDF & PCF";
                }
                catch
                { }
            }
        }

        private void manualToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                try
                {
                    row.Cells["Take_Off_Method"].Value = "Manual";
                }
                catch
                { }
            }
        }

        private void exportMTOAsSTOAndEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {

            DataTable dtFromGrid = new DataTable();
            dtFromGrid = (dataGridView1.DataSource as DataTable).Copy();
            DataSet ds = new DataSet();
            try
            {
                dtFromGrid.Columns.Remove("Material Code");
                dtFromGrid.Columns.Remove("Spool Number");
                dtFromGrid.Columns.Remove("recdate");
                dtFromGrid.Columns.Remove("linenum");
                dtFromGrid.Columns.Remove("revnum");
                dtFromGrid.Columns.Remove("linesize");
                dtFromGrid.Columns.Remove("Index");
                
                ds.Tables.Clear();
                ds.Tables.Add(dtFromGrid);

            }
            catch { }

            ExportDataSetToExcel(ds, jobnum + "_" + trans + "_STO");


            //Outlook.MailItem mailItem = (Outlook.MailItem)
            // this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            Outlook.Application app = new Outlook.Application();
            Outlook.MailItem mailItem = app.CreateItem(Outlook.OlItemType.olMailItem);
            mailItem.Subject = jobnum + "__" + trans + "_STO Report";
            mailItem.To = "kenneth.smith@EpicPiping.com; sunil.gawli@epicpiping.com; Shailesh.Dabhekar@epicpiping.com; Adam.Martin@epicpiping.com";
            mailItem.CC = "andre.naquin@EpicPiping.com; kevin.flores@epicpiping.com; Monty.Cornes@EpicPiping.com; travis.stromain@EpicPiping.com";
            //mailItem.Body = "Please see the attached document.";
            mailItem.Attachments.Add(stopath);//logPath is a string holding path to the log.txt file
            //mailItem.Importance = Outlook.OlImportance.olImportanceHigh;
            mailItem.Display(mailItem);
            mailItem.HTMLBody = "Please see the attached document." + mailItem.HTMLBody;
        }

        private void newToolStripMenuItem1_Click(object sender, EventArgs e)
        { 
            jobnum = Microsoft.VisualBasic.Interaction.InputBox("Job Number", "New Report Info", "Default");
            trans = Microsoft.VisualBasic.Interaction.InputBox("Transmittal", "New Report Info", "Default");

            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                InitialDirectory = @"V:\MTO\Spoolgen\Reports\Original_Reports\",
                Title = "Browse for CSV Report",
                CheckFileExists = true,
                CheckPathExists = true,

                DefaultExt = "CSV",
                Filter = "Csv Files (*.CSV)|*.csv",
                FilterIndex = 2,
                RestoreDirectory = true,
                ReadOnlyChecked = true,
                ShowReadOnly = true
            };
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                DataTable dataTable3 = new DataTable();
                dataTable3 = ConvertToDataTable(openFileDialog1.FileName, 1);
                dataGridView1.DataSource = dataTable3;

            }
            //foreach (DataRow row in dataGridView1.Rows)
            // {
            // this.GetDataBy()
            //}
            dataGridView1.Columns["Spool Number"].Visible = false;
            dataGridView1.Columns["Material Code"].Visible = false;
            dataGridView1.Columns["recdate"].Visible = false;
            dataGridView1.Columns["linenum"].Visible = false;
            dataGridView1.Columns["revnum"].Visible = false;
            dataGridView1.Columns["linesize"].Visible = false;
            dataGridView1.Columns["index"].Visible = false;
            dataGridView1.AutoResizeColumns();
            dataGridView2.AutoResizeColumns();
            String timeStamp = GetTimestamp(DateTime.Now);
            if (File.Exists(@"V:\MTO\Spoolgen\Reports\Original_Reports\Material_" + jobnum + "_" + trans + ".csv"))
            {
                File.Move(openFileDialog1.FileName, @"V:\MTO\Spoolgen\Reports\Original_Reports\Material_" + jobnum + "_" + trans + timeStamp + ".csv");
            }
            if (!File.Exists(@"V:\MTO\Spoolgen\Reports\Original_Reports\Material_" + jobnum + "_" + trans + ".csv"))
            {
                File.Move(openFileDialog1.FileName, @"V:\MTO\Spoolgen\Reports\Original_Reports\Material_" + jobnum + "_" + trans + ".csv");
            }


        }

        private void pcmkToolStripMenuItem4_Click(object sender, EventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtext = "";
                }
                catch { }
            }
        }

        private void toolStripMenuItem20_Click(object sender, EventArgs e)
        {
            string oldtext = "C-7900-PI"; //Microsoft.VisualBasic.Interaction.InputBox("Text to Replace", "Find and Replace - RefDwg", "Default");
            string newtext = Microsoft.VisualBasic.Interaction.InputBox("Replacing text with", "Find and Replace - RefDwg", "Default");



            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                try
                {
                    string replaceing = row.Cells["Piecemark"].Value.ToString();
                    string lastsegment = row.Cells["Piecemark"].Value.ToString().Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2);
                    string firstsegment = row.Cells["Piecemark"].Value.ToString().Substring(0, row.Cells["Piecemark"].Value.ToString().Length - 2);
                    string newtextletter = "";


                    switch (lastsegment)
                    {
                        case "-1":
                            newtextletter = "A";
                            break;
                        case "-2":
                            newtextletter = "B";
                            break;
                        case "-3":
                            newtextletter = "C";
                            break;
                        case "-4":
                            newtextletter = "D";
                            break;
                        case "-5":
                            newtextletter = "E";
                            break;
                        case "-6":
                            newtextletter = "F";
                            break;
                        case "-7":
                            newtextletter = "G";
                            break;
                        case "-8":
                            newtextletter = "H";
                            break;
                        case "-9":
                            newtextletter = "I";
                            break;
                    }
                            replaceing = row.Cells["Piecemark"].Value.ToString();
                    string TEMP = replaceing.Substring(row.Cells["Piecemark"].Value.ToString().Length - 2, 2).Replace(lastsegment, firstsegment + newtextletter);
                    row.Cells["Piecemark"].Value = TEMP.Replace(oldtext, newtext);
                }
                catch
                { }
            }
        }

        private void backgroundWorker1_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            //
             try
            {
            string fivedigitjobnum = dataGridView1.Rows[0].Cells["Production_No"].Value.ToString();
            DataTable SPOOLtable_table = isologchecker.GetDataBy1(fivedigitjobnum.Substring(fivedigitjobnum.Length - 5, 5));
            foreach (DataGridViewRow item in dataGridView1.Rows)
            {
                if (String.IsNullOrEmpty(item.Cells[0].Value as String))
                {
                    break;
                }
                if (item.Cells["Piecemark"].Value.ToString() != "")
                {

                    if (SPOOLtable_table != null)
                    {
                        string pcmk = item.Cells["Piecemark"].Value.ToString();
                        string find = "spool_pcmark = '" + pcmk + "'";

                        DataRow[] foundRows = SPOOLtable_table.Select(find);
                        int fr = foundRows.Count();
                        if (fr >= 1)
                        {

                            if (SPOOLtable_table.Rows[0]["spool"] != DBNull.Value)
                            {
                                item.Cells["Spool Number"].Value = foundRows[0]["spool"].ToString();
                            }
                            if (SPOOLtable_table.Rows[0]["isoLog_recvDate"].ToString() != null)
                            {
                                item.Cells["recdate"].Value = foundRows[0]["isoLog_recvDate"].ToString();
                            }
                            if (SPOOLtable_table.Rows[0]["isoLog_LineNum"].ToString() != null)
                            {
                                item.Cells["linenum"].Value = foundRows[0]["isoLog_LineNum"].ToString();
                            }
                            if (SPOOLtable_table.Rows[0]["isoLog_revNum"].ToString() != null)
                            {
                                item.Cells["revnum"].Value = foundRows[0]["isoLog_revNum"].ToString();
                            }
                            if (SPOOLtable_table.Rows[0]["isoLog_LineSize"].ToString() != null)
                            {
                                item.Cells["linesize"].Value = foundRows[0]["isoLog_LineSize"].ToString();
                            }
                        }

                    }

                }
                DataTable ISOLOGtable_table = isologchecker.GetData(item.Cells["Pipeline_Reference"].Value.ToString());
                if (ISOLOGtable_table.Rows.Count != 0)
                {
                    item.Cells["Material Code"].Value = ISOLOGtable_table.Rows[0]["isoLog_mat"].ToString();
                }
                if (ISOLOGtable_table.Rows.Count != 0)
                {
                    item.Cells["Source"].Value = "T" + ISOLOGtable_table.Rows[ISOLOGtable_table.Rows.Count - 1]["isoLog_transNum"].ToString();
                }
                if (ISOLOGtable_table.Rows.Count == 1)
                {
                    item.Cells["Source"].Value = "T" + ISOLOGtable_table.Rows[0]["isoLog_transNum"].ToString();
                }
                if (ISOLOGtable_table.Rows.Count >= 2)
                {
                    item.Cells["Source"].Value = "T" + ISOLOGtable_table.Rows[ISOLOGtable_table.Rows.Count - 1]["isoLog_transNum"].ToString() + "-REV";
                }

            }

            }

             catch { }
            //label3.Visible = false;
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            
        }

        private void backgroundWorker2_DoWork(object sender, System.ComponentModel.DoWorkEventArgs e)
        {
            
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, System.ComponentModel.RunWorkerCompletedEventArgs e)
        {
            label3.Visible = false;
        }

        public static String GetTimestamp(DateTime value)
        {
            return value.ToString("yyyyMMddHHmmssffff");
        }
    }
}
