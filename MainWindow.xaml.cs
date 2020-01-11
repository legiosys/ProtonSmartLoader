using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using Excel = Microsoft.Office.Interop.Excel;

namespace Proton_Loader
{

    public partial class MainWindow : Window
    {
        OpenFileDialog ofd = new OpenFileDialog();
        SaveFileDialog sfd = new SaveFileDialog();
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel._Worksheet xlWorksheet;
        Dictionary<string,int> headers = new Dictionary<string,int>();
        DataTable sheet;
        public MainWindow()
        {
            InitializeComponent();
            loadTemplates();
        }
        private class Templ //класс для комбобокса шаблона
        {
            public string name;
            public string id;
            public Templ (string n, string i)
            {
                name = n;
                id = i;
            }
            public override string ToString()
            {
                return name;
            }
        }
        private class ProgressClass //класс для отправки прогресса загрузки
        {
            public int percent;
            public string text;
        }
        private void loadTemplates() //загрузка списка шаблонов из базы ProtonSmartPrint (sw_database, может быть другим)
        {
            try
            {
                DataTable dt = new DataTable();
                SQLiteConnection conn = new SQLiteConnection("Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                    "\\Geksagon\\ProtonSmartPrint Управление\\sw_database.sqlite;Version=3;");
                conn.Open();
                SQLiteDataAdapter adapter = new SQLiteDataAdapter("select name,ext_id from templates",conn);
                adapter.Fill(dt);
                foreach(DataRow row in dt.Rows)
                    cb_templates.Items.Add(new Templ((string)row["name"], (string)row["ext_id"]));
                cb_templates.SelectedIndex = 0;
                conn.Close();
            }
            catch (SQLiteException) {}
        }
        //кнопка обработать
        private async void Button_Click(object sender, RoutedEventArgs e)
        {
            // проверки на заполнение
            if(lbl_open.Content.Equals("") || lbl_save.Content.Equals(""))
            {
                MessageBox.Show("Не выбран файл!","Ошибка");
                return;
            }
            if (cb_barcode.SelectedIndex < 0 || cb_id.SelectedIndex < 0 || cb_name.SelectedIndex < 0)
            {
                MessageBox.Show("Не выбраны поля!", "Ошибка");
                return;
            }
            Templ temp = ((Templ)cb_templates.SelectedItem);
            var progress = new Progress<ProgressClass>(progr => //конструкция для асинк прогресс бара
            {                
                pb_loading.Value = progr.percent;
                lbl_state.Content = progr.text;
                if (progr.percent == 100) pb_loading.Maximum = 99;
            });
            string name = cb_name.Text;
            string id = cb_id.Text;
            string barcode = cb_barcode.Text;
            await Task.Run(() =>
            {
                LoadWorksheet(temp, progress, name, id, barcode);
            });
        }
        //обработка Excel файла
        void LoadWorksheet(Templ temp, IProgress<ProgressClass> progress, string name, string id, string barcode) //загрузка из эксель в тхт
        {
            ProgressClass pr = new ProgressClass();
            int rowIndex = 2;
            int i=1;
            DataRow row;
            while ((xlWorksheet.Cells[i*50, 1]).Value2 != null) //подсчет примерного количества записей (для прогрессбара)
            {
                i++;
            }
            i = i * 50;
            string outrow;               
            using (System.IO.StreamWriter file = new System.IO.StreamWriter(sfd.FileName)) //открываем на запись
            {
                while ((xlWorksheet.Cells[rowIndex, 1]).Value2 != null)
                {
                    pr.percent = Convert.ToInt32(100 * rowIndex / i); //процент загрузки
                    pr.text = "Обработка данных...";
                    progress.Report(pr);
                    row = sheet.NewRow();
                    foreach (var key in headers.Keys)//получение строки excel
                        row[key] = xlWorksheet.Cells[rowIndex, headers[key]].Value2;
                    rowIndex++; 
                    string fcolumns = "";
                    foreach (var key in headers.Keys) //формирование пользовательских полей для бд
                        if (!key.Equals(name) && !key.Equals(barcode))
                            fcolumns = string.Format("{0}{1};", fcolumns, row[key]);
                    for (int j = 0; j < 30 - headers.Count+2; j++)
                        fcolumns += ";"; //сборка строки csv-файла
                    outrow = string.Format("{0};0;;;;{1}{2};{3};;0;0;0;0;0;3;3;0;{4};0;1;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;",
                       row[name], fcolumns, row[id], temp.name, row[barcode]);
                    file.WriteLine(outrow);
                }
            }
            pr.percent = Convert.ToInt32(100);
            pr.text = "Обработано " + (rowIndex-2) + " записей.";
            progress.Report(pr);
        }
        //выбор исходного Excel файла
        private void btn_open_Click(object sender, RoutedEventArgs e) //диалог открытия
        {
            ofd.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx";
            ofd.ShowDialog();
            lbl_open.Content = ofd.FileName;
            xlApp = new Excel.Application();
            xlWorkbook = xlApp.Workbooks.Open(ofd.FileName);
            xlWorksheet = xlWorkbook.Sheets[1];
            int i = 1;
            //создание словаря заголовков столбцов 
            while ((xlWorksheet.Cells[1, i]).Value2 != null)
                headers.Add(xlWorksheet.Cells[1, i].Value2,i++);
            //подстановка соурса в комбобоксы и выбор(если имена столбцов по-умолчанию) необходимых значений
            cb_name.ItemsSource = headers.Keys;
            cb_name.SelectedItem = "Наименование";
            cb_id.ItemsSource = headers.Keys;
            cb_id.SelectedItem = "Idtovar";
            cb_barcode.ItemsSource = headers.Keys;
            cb_barcode.SelectedItem = "Barcode";
            sheet = new DataTable();
            foreach (var key in headers.Keys) //создание стоблцов
                sheet.Columns.Add(key);
        }

        private void btn_save_Click(object sender, RoutedEventArgs e) //диалог сохранения
        {
            sfd.Filter = "Text File|*.txt";
            sfd.DefaultExt = ".txt";
            sfd.ShowDialog();
            lbl_save.Content = sfd.FileName;
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (xlWorksheet != null)//оcвобождение памяти от Excel
                Marshal.ReleaseComObject(xlWorksheet);
            if (xlWorkbook != null)
            {
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);               
            }
            if (xlApp != null)
            {
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }                
        
        }
    }
}
