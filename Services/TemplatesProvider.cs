using Proton_Loader.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;

namespace Proton_Loader.Services
{
    public class TemplatesProvider
    {
        public TemplatesProvider()
        {
            Templates = new List<Template>();
        }
        public List<Template> Templates { get; }
        public void Load() //загрузка списка шаблонов из базы ProtonSmartPrint (sw_database, может быть другим)
        {
            try
            {
                DataTable dt = new DataTable();
                SQLiteConnection conn = new SQLiteConnection("Data Source=" + Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) +
                    "\\Geksagon\\ProtonSmartPrint Управление\\sw_database.sqlite;Version=3;");
                conn.Open();
                SQLiteDataAdapter adapter = new SQLiteDataAdapter("select name,ext_id from templates", conn);
                adapter.Fill(dt);
                foreach (DataRow row in dt.Rows)
                    Templates.Add(new Template() { Id = (string)row["ext_id"], Name = (string)row["name"] });
                conn.Close();
            }
            catch (SQLiteException ex)
            {
                MessageBox.Show($"Произошла ошибка базы данных, возможно файл перенесен\n {ex.Message}");
            }
        }

        public void FillComboBox(ref ComboBox comboBox)
        {
            comboBox.ItemsSource = Templates;
            comboBox.SelectedIndex = 0;
        }
    }
}
