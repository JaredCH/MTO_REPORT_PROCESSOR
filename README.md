# MTO_REPORT_PROCESSOR



        private void GetExternalData_Top2Bottom()
        {
            string fivedigitjobnum_threaded = dataGridView1.Rows[0].Cells["Production_No"].Value.ToString();
            DataTable SPOOLtable_table_threaded = isologchecker.GetDataBy1(fivedigitjobnum.Substring(fivedigitjobnum.Length - 5, 5));
            for (int i = 0; i < dataGridView1.RowCount / 2; i++)
            {
                string pcmk_threaded = dataGridView1.Rows[i].Cells["Piecemark"].Value.ToString();
                string find_threaded = "spool_pcmark = '" + pcmk_threaded + "'";
                DataRow[] foundRows_threaded = SPOOLtable_table_threaded.Select(find);
                int fr = foundRows_threaded.Count();
                if (fr >= 1)
                {

                    if (SPOOLtable_table_threaded.Rows[0]["spool"] != DBNull.Value)
                    {
                        dataGridView1.Rows[i].Cells["Spool Number"].Value = foundRows_threaded[0]["spool"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_recvDate"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["recdate"].Value = foundRows_threaded[0]["isoLog_recvDate"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_LineNum"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["linenum"].Value = foundRows_threaded[0]["isoLog_LineNum"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_revNum"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["revnum"].Value = foundRows_threaded[0]["isoLog_revNum"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_LineSize"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["linesize"].Value = foundRows_threaded[0]["isoLog_LineSize"].ToString();
                    }
                }
            }
        }





        private void GetExternalData_Bottom2Top()
        {
            string fivedigitjobnum_threaded = dataGridView1.Rows[0].Cells["Production_No"].Value.ToString();
            DataTable SPOOLtable_table_threaded = isologchecker.GetDataBy1(fivedigitjobnum_threaded.Substring(fivedigitjobnum_threaded.Length - 5, 5));
            for (int i = dataGridView1.RowCount; i > dataGridView1.RowCount / 2; i++)
            {
                string pcmk_threaded = dataGridView1.Rows[i].Cells["Piecemark"].Value.ToString();
                string find_threaded = "spool_pcmark = '" + pcmk_threaded + "'";
                DataRow[] foundRows_threaded = SPOOLtable_table_threaded.Select(find_threaded);
                int fr = foundRows_threaded.Count();
                if (fr >= 1)
                {

                    if (SPOOLtable_table_threaded.Rows[0]["spool"] != DBNull.Value)
                    {
                        dataGridView1.Rows[i].Cells["Spool Number"].Value = foundRows_threaded[0]["spool"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_recvDate"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["recdate"].Value = foundRows_threaded[0]["isoLog_recvDate"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_LineNum"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["linenum"].Value = foundRows_threaded[0]["isoLog_LineNum"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_revNum"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["revnum"].Value = foundRows_threaded[0]["isoLog_revNum"].ToString();
                    }
                    if (SPOOLtable_table_threaded.Rows[0]["isoLog_LineSize"].ToString() != null)
                    {
                        dataGridView1.Rows[i].Cells["linesize"].Value = foundRows_threaded[0]["isoLog_LineSize"].ToString();
                    }
                }
            }
        }
