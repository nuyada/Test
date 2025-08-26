using System;
using System.Data;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using DataTable = System.Data.DataTable;


namespace Test
{
    public partial class MainForm : Form
    {
        private DataTable dataTable;

        public MainForm()
        {
            InitializeComponent();
            InitializeDataTable();
            InitializeDataGridView();
        }

        private void InitializeDataTable()
        {
            dataTable = new DataTable();
            dataTable.Columns.Add("ID", typeof(int));
            dataTable.Columns.Add("Name", typeof(string));
            dataTable.Columns.Add("Email", typeof(string));
            dataTable.Columns.Add("Phone", typeof(string));
            dataTable.Columns.Add("Salary", typeof(decimal));
        }

        private void InitializeDataGridView()
        {
            dataGridView.DataSource = dataTable;
            dataGridView.AutoGenerateColumns = true;
            dataGridView.AllowUserToAddRows = true;
            dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView.MultiSelect = true;
        }

        // Импорт данных из Excel
        private void btnImport_Click(object sender, EventArgs e)
        {
            try
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                    openFileDialog.Title = "Выберите Excel файл";

                    if (openFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ImportExcelData(openFileDialog.FileName);
                        statusLabel.Text = $"Данные импортированы из: {Path.GetFileName(openFileDialog.FileName)}";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при импорте: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ImportExcelData(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = workbook.Sheets[1];
                range = worksheet.UsedRange;

                dataTable.Clear();

                
                for (int col = 1; col <= range.Columns.Count; col++)
                {
                    string columnName = (range.Cells[1, col] as Range)?.Value2?.ToString() ?? $"Column{col}";
                    if (!dataTable.Columns.Contains(columnName))
                    {
                        dataTable.Columns.Add(columnName);
                    }
                }

                
                for (int row = 2; row <= range.Rows.Count; row++)
                {
                    DataRow dataRow = dataTable.NewRow();
                    for (int col = 1; col <= range.Columns.Count; col++)
                    {
                        if (range.Cells[row, col] != null &&
                            (range.Cells[row, col] as Range)?.Value2 != null)
                        {
                            dataRow[col - 1] = (range.Cells[row, col] as Range).Value2.ToString();
                        }
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            finally
            {
                // Освобождаем ресурсы
                if (range != null) Marshal.ReleaseComObject(range);
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Экспорт данных в Excel
        private void btnExport_Click(object sender, EventArgs e)
        {
            try
            {
                using (SaveFileDialog saveFileDialog = new SaveFileDialog())
                {
                    saveFileDialog.Filter = "Excel Files|*.xlsx";
                    saveFileDialog.Title = "Сохранить как Excel файл";

                    if (saveFileDialog.ShowDialog() == DialogResult.OK)
                    {
                        ExportToExcel(saveFileDialog.FileName);
                        statusLabel.Text = $"Данные экспортированы в: {Path.GetFileName(saveFileDialog.FileName)}";
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ExportToExcel(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = null;
            Workbook workbook = null;
            Worksheet worksheet = null;

            try
            {
                excelApp = new Microsoft.Office.Interop.Excel.Application();
                workbook = excelApp.Workbooks.Add();
                worksheet = workbook.Sheets[1];

                // Записываем заголовки
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1] = dataTable.Columns[i].ColumnName;
                }

                // Записываем данные
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1] = dataTable.Rows[i][j].ToString();
                    }
                }

                // Форматируем заголовки
                Range headerRange = worksheet.Range[worksheet.Cells[1, 1],
                    worksheet.Cells[1, dataTable.Columns.Count]];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(
                    System.Drawing.Color.LightGray);

                // Автоподбор ширины колонок
                worksheet.Columns.AutoFit();

                workbook.SaveAs(filePath);
                MessageBox.Show("Данные успешно экспортированы!", "Успех",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                // Освобождаем ресурсы
                if (worksheet != null) Marshal.ReleaseComObject(worksheet);
                if (workbook != null)
                {
                    workbook.Close();
                    Marshal.ReleaseComObject(workbook);
                }
                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        // Добавление новой строки
        private void btnAddRow_Click(object sender, EventArgs e)
        {
            try
            {
                DataRow newRow = dataTable.NewRow();
                dataTable.Rows.Add(newRow);
                statusLabel.Text = "Добавлена новая строка";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при добавлении строки: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Удаление выбранных строк
        private void btnDeleteRows_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView.SelectedRows.Count > 0)
                {
                    DialogResult result = MessageBox.Show(
                        $"Вы уверены, что хотите удалить {dataGridView.SelectedRows.Count} строк?",
                        "Подтверждение удаления",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question);

                    if (result == DialogResult.Yes)
                    {
                        foreach (DataGridViewRow row in dataGridView.SelectedRows)
                        {
                            if (!row.IsNewRow)
                            {
                                dataGridView.Rows.Remove(row);
                            }
                        }
                        statusLabel.Text = $"Удалено {dataGridView.SelectedRows.Count} строк";
                    }
                }
                else
                {
                    MessageBox.Show("Выберите строки для удаления", "Информация",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при удалении строк: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Сохранение изменений
        private void btnSaveChanges_Click(object sender, EventArgs e)
        {
            try
            {
                dataGridView.EndEdit();
                statusLabel.Text = "Изменения сохранены";
                MessageBox.Show("Все изменения сохранены успешно!", "Успех",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении изменений: {ex.Message}", "Ошибка",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            statusLabel.Text = "Готов к работе. Загрузите данные из Excel или начните редактирование.";
        }
    }
}
