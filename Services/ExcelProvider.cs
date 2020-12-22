using OfficeOpenXml;
using Proton_Loader.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;

namespace Proton_Loader.Services
{
    public class ExcelProvider
    {
        private ExcelPackage _excel;
        private ExcelWorksheet _sheet => _excel.Workbook.Worksheets.First();
        public int RowsCount => _sheet.Dimension.End.Row;
        public int ColumnsCount => _sheet.Dimension.End.Column;
        public ExcelProvider(string path)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            FileInfo existingFile = new FileInfo(path);
            _excel = new ExcelPackage(existingFile);
        }

        public Dictionary<int, string> GetHeaders()
            => GetRow(1);

        public async Task ProcessFile(IProgress<ProgressBarState> progress, Template templ, int name, int id, int barcode, string file)
        {
            var pr = new ProgressBarState();
            var writer = new StreamWriter(file, false, Encoding.UTF8);
            var headers = GetHeaders().Where(x => x.Key != name && x.Key != barcode); //пользовательские поля
            for (int rowIndex = 2; rowIndex <= RowsCount; rowIndex++)
            {
                pr.Percent = Convert.ToInt32(100 * RowsCount / rowIndex); //процент загрузки
                pr.Text = "Обработка данных...";
                progress.Report(pr);
                var row = GetRow(rowIndex);
                var userFields = BuildUserFields(row.Where(x => headers.Any(y => y.Key == x.Key)));
                var outRow = $"{_sheet.Cells[rowIndex, name].Value};0;;;;{userFields};{row[id]};{templ.Name};;0;0;0;0;0;3;3;0;{row[barcode]};0;1;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;";
                await writer.WriteLineAsync(outRow);
            }
            await writer.FlushAsync();
            pr.Percent = Convert.ToInt32(100);
            pr.Text = "Обработано " + (RowsCount - 1) + " записей.";
            progress.Report(pr);
        }

        private string BuildUserFields(IEnumerable<KeyValuePair<int,string>> row)
        {
            var result = "";
            foreach (var cell in row)
            {
                result += $"{cell.Value};";
            }
            for (int j = 0; j < 30 - ColumnsCount + 2; j++)
                result += ";";

            return result;
        }

        private Dictionary<int,string> GetRow(int rowIndex)
        {
            var result = new Dictionary<int, string>();
            for (int c = 1; c <= ColumnsCount; c++)
                result.Add(c, _sheet.Cells[rowIndex, c].Value.ToString().Trim());
            return result;
        }
    }
}
