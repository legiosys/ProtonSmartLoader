using Microsoft.Win32;
using Proton_Loader.Models;
using Proton_Loader.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;

namespace Proton_Loader
{

    public partial class MainWindow : Window
    {
        OpenFileDialog ofd = new OpenFileDialog();
        SaveFileDialog sfd = new SaveFileDialog();

        private ExcelProvider _excelProvider;
        public MainWindow()
        {
            InitializeComponent();
            var templateProvider = new TemplatesProvider();
            templateProvider.Load();
            templateProvider.FillComboBox(ref cb_templates);
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
            var temp = (Template)cb_templates.SelectedItem;
            var progress = new Progress<ProgressBarState>(progr => //конструкция для асинк прогресс бара
            {                
                pb_loading.Value = progr.Percent;
                lbl_state.Content = progr.Text;
                if (progr.Percent == 100) pb_loading.Maximum = 99;
            });

            var name = (KeyValuePair<int, string>)cb_name.SelectedItem;
            var id = (KeyValuePair<int, string>)cb_id.SelectedItem;
            var barcode = (KeyValuePair<int, string>)cb_barcode.SelectedItem;

            await _excelProvider.ProcessFile(progress, temp, name.Key, id.Key, barcode.Key, sfd.FileName);
        }

        //выбор исходного Excel файла
        private void btn_open_Click(object sender, RoutedEventArgs e) //диалог открытия
        {
            ofd.Filter = "Excel Files(*.xls;*.xlsx)|*.xls;*.xlsx";
            ofd.ShowDialog();
            lbl_open.Content = "Загрузка файла...";

            _excelProvider = new ExcelProvider(ofd.FileName);
            var headers = _excelProvider.GetHeaders();
            lbl_open.Content = ofd.FileName;
            //подстановка соурса в комбобоксы и выбор(если имена столбцов по-умолчанию) необходимых значений
            cb_name.ItemsSource = headers;
            cb_name.SelectedItem = headers.FirstOrDefault(x => x.Value.Equals("Наименование"));
            cb_id.ItemsSource = headers;
            cb_id.SelectedItem = headers.FirstOrDefault(x => x.Value.Equals("Idtovar")); 
            cb_barcode.ItemsSource = headers;
            cb_barcode.SelectedItem = headers.FirstOrDefault(x => x.Value.Equals("Barcode"));
        }

        private void btn_save_Click(object sender, RoutedEventArgs e) //диалог сохранения
        {
            sfd.Filter = "Text File|*.txt";
            sfd.DefaultExt = ".txt";
            sfd.ShowDialog();
            lbl_save.Content = sfd.FileName;
        }
    }
}
