using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace debet_kredit_xls
{
    public partial class Form1 : Form
    {
        DataSet ds = new DataSet();
        DataTable dt = new DataTable();
        
        public Form1()
        {
            InitializeComponent();
            dt.Columns.Add("Статус клиента", Type.GetType("System.String"));
            dt.Columns.Add("№ договора", Type.GetType("System.String"));
            dt.Columns.Add("Баланс", Type.GetType("System.Double"));

        }

        private void добавитьДанныеИзExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog o = new OpenFileDialog();
            o.Filter = "excel files (*.xls)|*.xls|All files (*.*)|*.*";
            o.RestoreDirectory = true;
            string file = string.Empty;
            if (o.ShowDialog() == DialogResult.OK)
            {
                file = o.FileName;
                if (File.Exists(file))
                {
                    LoadDataFromFile(file);
                }
            }
        }

        private void LoadDataFromFile(string file)
        {
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync(file);
            }

            
        }


        private void очитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dt.Rows.Clear();
            dataGridView1.Rows.Clear();
        }

        private void сохранитьОтсортированныеДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog s = new SaveFileDialog();
            s.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            s.RestoreDirectory = true;
            string file = string.Empty;
            if (s.ShowDialog() == DialogResult.OK)
            {
                SaveDataToFile(s.FileName);
                MessageBox.Show("Данные сохранены");
            }
        }

        private void SaveDataToFile(string file)
        {
            string scv_data = string.Empty;
            //for (int f = 0; f < dataGridView1.ColumnCount; f++)
            //{
            //    scv_data += (dataGridView1.Columns[f].HeaderText + ";");

            //}
            //scv_data += "\t\n";
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                int flag = 0;
                if (dataGridView1.Rows[i].Cells[0].Value.ToString() == "Информирование")
                {
                    flag = 1;
                }
               scv_data += string.Format("{0};{0};{1};1000;{2};{3}",dataGridView1.Rows[i].Cells[1].Value.ToString(), dataGridView1.Rows[i].Cells[2].Value.ToString(), DateTime.Now.ToString("dd-MM-yyyy"),flag);
               scv_data += "\t\n";
            }
            StreamWriter wr = new StreamWriter(file, false, Encoding.GetEncoding("windows-1251"));
            wr.Write(scv_data);
            wr.Close();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string file = (string)e.Argument;
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application(); //Excel
                Microsoft.Office.Interop.Excel.Workbook xlWB; //рабочая книга              
                Microsoft.Office.Interop.Excel.Worksheet xlSht; //лист Excel   
                xlWB = xlApp.Workbooks.Open(file); //название файла Excel                                             
                xlSht = xlWB.Worksheets["Sheet1"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
                var arrData = (object[,])xlSht.Range["A3:G" + iLastRow].Value; //берём данные с листа Excel
                                                                               //xlApp.Visible = true; //отображаем Excel     
                xlWB.Close(false); //закрываем книгу, изменения не сохраняем
                xlApp.Quit(); //закрываем Excel

                e.Result = arrData;
            }
            catch (SystemException ex)
            {
                e.Result = null;
            }
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null && e.Result != null)
            {
                var arrData = (object[,])e.Result;
                dt.Rows.Clear();
                int RowsCount = arrData.GetUpperBound(0);
                int ColumnsCount = arrData.GetUpperBound(1);
                //заполняем DataGridView данными из массива
                for (int i = 1; i <= RowsCount; i++)
                {
                    DataRow dr = dt.NewRow();
                    dr[1] = arrData[i, 1];
                    double balans = 0;
                    if (arrData[i, 6] != null)
                    {
                        try
                        {
                            balans = Convert.ToDouble(arrData[i, 6]);

                        }
                        catch (SystemException ex)
                        {
                            dr[2] = -1;
                            dr[0] = "ошибка исходных данных";
                            dt.Rows.Add(dr);
                            continue;
                        }
                    }
                    else if (arrData[i, 7] != null)
                    {
                        try
                        {
                            balans = -Convert.ToDouble(arrData[i, 7]);
                        }
                        catch (SystemException ex)
                        {
                            balans = -1;
                            dr[0] = "ошибка исходных данных";
                            dt.Rows.Add(dr);
                            continue;
                        }
                    }

                    dr[2] = balans;

                    if (balans < 0)
                    {
                        dr[0] = "Информирование";
                    }
                    else
                    {
                        dr[0] = "Нет задолженности";
                    }

                    dt.Rows.Add(dr);


                }

                dataGridView1.DataSource = dt;
                FilterData();
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FilterData();

        }

        private void FilterData()
        {
            if (comboBox1.Text == "Все")
            {
                dataGridView1.DataSource = dt;
            }
            else if (comboBox1.Text == "Должники")
            {
                DataTable dt1 = dt.Copy();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (dt1.Rows[i]["Статус клиента"].ToString() != "Информирование")
                    {
                        dt1.Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                dataGridView1.DataSource = dt1;
                //DataRow[] rows = dt.Select("[Статус клиента] = 'Информирование'");

            }
            else if (comboBox1.Text == "Без долга")
            {
                //DataRow[] rows = dt.Select("[Статус клиента] = 'Нет задолженности'");
                DataTable dt1 = dt.Copy();
                for (int i = 0; i < dt1.Rows.Count; i++)
                {
                    if (dt1.Rows[i]["Статус клиента"].ToString() != "Нет задолженности")
                    {
                        dt1.Rows.RemoveAt(i);
                        i--;
                        continue;
                    }
                }

                dataGridView1.DataSource = dt1;
            }
        }
    }
}
