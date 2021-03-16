using System;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace debet_kredit_xls
{
    public partial class Form1 : Form
    {
        /// <summary>
        /// Таблица с данными
        /// </summary>
        DataTable dataTable = new DataTable();
        
        /// <summary>
        /// Конструктор формы
        /// </summary>
        public Form1()
        {
            InitializeComponent();
            dataTable.Columns.Add("Статус клиента", Type.GetType("System.String"));
            dataTable.Columns.Add("№ договора", Type.GetType("System.String"));
            dataTable.Columns.Add("Баланс", Type.GetType("System.Double"));
            
        }

        /// <summary>
        /// Обработчик события клика по кнопке "Добавить файл..."
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void addDataFromExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog();
            fileDialog.Filter = "excel files (*.xls)|*.xls|All files (*.*)|*.*";
            fileDialog.RestoreDirectory = true;
            fileDialog.InitialDirectory = AppContext.BaseDirectory;
            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                string file = fileDialog.FileName;
                if (File.Exists(file))
                {
                    LoadDataFromFile(file);
                }
            }
        }

        /// <summary>
        /// Метод для старта воркера для считывания данных в фоновом режиме
        /// </summary>
        /// <param name="file"></param>
        private void LoadDataFromFile(string file)
        {
            if (!backgroundWorker1.IsBusy)
            {
                backgroundWorker1.RunWorkerAsync(file);
            }            
        }

        /// <summary>
        /// Обработчик события клика по кнопке "Очистить данные"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void clearDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (dataTable.Rows.Count > 0)
                dataTable.Rows.Clear();
        }

        /// <summary>
        /// Обработчик события на клик по кнопке "Сохранить данные"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void saveDataToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "csv files (*.csv)|*.csv|All files (*.*)|*.*";
            saveDialog.RestoreDirectory = true;
            saveDialog.InitialDirectory = AppContext.BaseDirectory;
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                SaveDataToFile(saveDialog.FileName);
                MessageBox.Show("Данные сохранены");
            }
        }

        /// <summary>
        /// Метод заполняет транспортный файл и сохраняет его в локальную директорию
        /// </summary>
        /// <param name="filePath"></param>
        private void SaveDataToFile(string filePath)
        {
            string csvData = string.Empty;
            string currentDate = DateTime.Now.ToString("dd-MM-yyyy");
            for (int i = 0; i < dataGridView1.RowCount - 1; i++)
            {
                int flag = 0;
                string status = dataGridView1.Rows[i].Cells[0].Value.ToString();
                string documentNumber = dataGridView1.Rows[i].Cells[1].Value.ToString();
                string balance = dataGridView1.Rows[i].Cells[2].Value.ToString();
                switch (status)
                {
                    case "Информирование":
                        flag = 1;
                        break;
                    case "Предупреждение":
                        flag = 2;
                        break;
                    case "Ограничение функций":
                        flag = 3;
                        break;
                }

                //Формат транспортного файла 1;2;3;4;5;6, где
                // 1 - Номер объекта
                // 2 - Номер договора
                // 3 - Баланс
                // 4 - Абонентская плата
                // 5 - Дата списания
                // 6 - Уровень информирования

                csvData += $"{documentNumber};{documentNumber};{balance};;{currentDate};{flag}";
                csvData += "\t\n";
            }
            StreamWriter wr = new StreamWriter(filePath, false, Encoding.GetEncoding("windows-1251"));
            wr.Write(csvData);
            wr.Close();
        }

        /// <summary>
        /// Метод считывает данные из Excel-файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            string fileName = (string)e.Argument;
            try
            {
                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application(); //Excel
                Microsoft.Office.Interop.Excel.Workbook xlWB; //рабочая книга              
                Microsoft.Office.Interop.Excel.Worksheet xlSht; //лист Excel   
                xlWB = xlApp.Workbooks.Open(fileName); //название файла Excel                                             
                xlSht = xlWB.Worksheets["Sheet1"]; //название листа или 1-й лист в книге xlSht = xlWB.Worksheets[1];
                int iLastRow = xlSht.Cells[xlSht.Rows.Count, "A"].End[Microsoft.Office.Interop.Excel.XlDirection.xlUp].Row;  //последняя заполненная строка в столбце А            
                var arrData = (object[,])xlSht.Range["A3:G" + iLastRow].Value; //берём данные с листа Excel
                                                                               //xlApp.Visible = true; //отображаем Excel     
                xlWB.Close(false); //закрываем книгу, изменения не сохраняем
                xlApp.Quit(); //закрываем Excel

                e.Result = arrData;
            }
            catch (Exception ex)
            {
                e.Result = null;
            }
        }

        /// <summary>
        /// Обработчик события об окончании считывания Excel-файла
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null || e.Result == null)
                return;

            var arrData = (object[,])e.Result;
            dataTable.Rows.Clear();
            int RowsCount = arrData.GetUpperBound(0);

            //заполняем DataGridView данными из массива
            for (int i = 1; i <= RowsCount; i++)
            {
                DataRow dataRow = dataTable.NewRow();
                dataRow[1] = arrData[i, 1];
                double balance = 0;
                if (arrData[i, 6] != null) //Дебет на конец периода
                {
                    try
                    {
                        balance = Convert.ToDouble(arrData[i, 6]);
                    }
                    catch
                    {
                        dataRow[2] = -1;
                        dataRow[0] = "Ошибка исходных данных";
                        dataTable.Rows.Add(dataRow);
                        continue;
                    }
                }
                else if (arrData[i, 7] != null) //Кредит на конец периода
                {
                    try
                    {
                        balance = -Convert.ToDouble(arrData[i, 7]);
                    }
                    catch
                    {
                        balance = -1;
                        dataRow[0] = "Ошибка исходных данных";
                        dataTable.Rows.Add(dataRow);
                        continue;
                    }
                }

                dataRow[2] = balance;

                if (balance < 0)
                {
                    dataRow[0] = "Информирование";
                    if (balance < -8000)
                        dataRow[0] = "Предупреждение";
                    if (balance < -15000)
                        dataRow[0] = "Ограничение функций";
                }
                else
                {
                    dataRow[0] = "Нет задолженности";
                }

                dataTable.Rows.Add(dataRow);
            }

            dataGridView1.DataSource = dataTable;
            //Расширяем колонку
            dataGridView1.Columns[0].Width = 150;
            FilterData();
        }

        /// <summary>
        /// Обработчик события при смене фильтра
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            FilterData();
        }

        /// <summary>
        /// Метод для фильтрации данных в таблице
        /// </summary>
        private void FilterData()
        {
            string status = "";
            switch (comboBox1.Text)
            {
                case "Должники - информирование":
                    status = "Информирование";
                    break;
                case "Должники - предупреждение":
                    status = "Предупреждение";
                    break;
                case "Должники - ограничение":
                    status = "Ограничение функций";
                    break;
                case "Без долга":
                    status = "Нет задолженности";
                    break;
                default:
                    dataGridView1.DataSource = dataTable;
                    return;
            }
            DataTable dataTableCopy = dataTable.Copy();
            for (int i = 0; i < dataTableCopy.Rows.Count; i++)
            {
                if (dataTableCopy.Rows[i]["Статус клиента"].ToString() != status)
                {
                    dataTableCopy.Rows.RemoveAt(i);
                    i--;
                    continue;
                }
            }

            dataGridView1.DataSource = dataTableCopy;
        }
    }
}
