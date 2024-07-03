using System;
using System.Data;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using ClosedXML.Excel;
using Fiscal.Classe;
using System.ComponentModel;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;

namespace Desafio2 {
    public partial class Form4 : Form {
        public Form4() {
            InitializeComponent();

            ContextMenuStrip contextMenu = new ContextMenuStrip();

            ToolStripMenuItem exportToExcel = new ToolStripMenuItem("Extrair para Excel");
            exportToExcel.Click += ExportToExcel_Click;
            contextMenu.Items.Add(exportToExcel);

            ToolStripMenuItem importFromExcel = new ToolStripMenuItem("Importar de Excel");
            importFromExcel.Click += ImportFromExcel_Click;
            contextMenu.Items.Add(importFromExcel);

            dataGridView1.ContextMenuStrip = contextMenu;

            LoadDataAsync();
        }

        public async Task LoadDataAsync() {
            try {
                using (var contexto = new DataContext.Contexto()) {
                    await contexto.Database.OpenConnectionAsync();

                    string selectSql = "SELECT * FROM TFornecedor";

                    using (FbCommand cmd = new FbCommand(selectSql, (FbConnection)contexto.Database.GetDbConnection())) {
                        using (FbDataReader reader = (FbDataReader)await cmd.ExecuteReaderAsync()) {
                            DataTable dataTable = new DataTable();
                            dataTable.Load(reader);

                            dataGridView1.DataSource = dataTable;

                            foreach (DataGridViewColumn column in dataGridView1.Columns) {
                                column.ReadOnly = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show($"Erro ao carregar dados: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void LoadData() {
            await LoadDataAsync();
        }

        private void ExportToExcel_Click(object sender, EventArgs e) {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (var workbook = new XLWorkbook()) {
                            var worksheet = workbook.Worksheets.Add("Sheet1");

                            for (int i = 0; i < dataGridView1.Columns.Count; i++) {
                                worksheet.Cell(1, i + 1).Value = dataGridView1.Columns[i].HeaderText;
                            }

                            for (int i = 0; i < dataGridView1.Rows.Count; i++) {
                                for (int j = 0; j < dataGridView1.Columns.Count; j++) {
                                    worksheet.Cell(i + 2, j + 1).Value = dataGridView1.Rows[i].Cells[j].Value?.ToString();
                                }
                            }

                            workbook.SaveAs(sfd.FileName);
                        }
                        MessageBox.Show("Exportação bem-sucedida!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Erro ao exportar para Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ImportFromExcel_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        using (var workbook = new XLWorkbook(ofd.FileName)) {
                            var worksheet = workbook.Worksheet(1);
                            var dataTable = dataGridView1.DataSource as DataTable ?? new DataTable();

                            var columnMapping = new Dictionary<string, string>
                            {                        
                                { "RAZAOSOCIAL", "RAZAOSOCIAL" },
                                { "NOMEFANTASIA", "NOMEFANTASIA" },
                                { "CNPJ", "CNPJ" },
                                { "CPF", "CPF" },
                                { "RG", "RG" },
                                { "IE", "IE" },
                                { "ENDERECO", "ENDERECO" },
                                { "COMPLEMENTO", "COMPLEMENTO" },
                                { "BAIRRO", "BAIRRO" },
                                { "CEP", "CEP" },
                                { "NUMERO", "NUMERO" },
                                { "CIDADE", "CIDADE" },
                                { "UF", "UF" },
                                { "TELEFONE", "TELEFONE" },
                                { "CELULAR", "CELULAR" },
                                { "EMAIL", "EMAIL" }
                            };

                            bool firstRow = true;
                            foreach (var row in worksheet.RowsUsed()) {
                                if (firstRow) {
                                    firstRow = false;
                                    continue;
                                }

                                bool exists = false;
                                foreach (DataRow existingRow in dataTable.Rows) {
                                    bool allMatch = true;
                                    foreach (var mapping in columnMapping) {
                                        var columnName = mapping.Value;
                                        var columnIndex = GetColumnIndex(worksheet, mapping.Key);
                                        var cellValue = row.Cell(columnIndex).Value.ToString();
                                        if (existingRow[columnName].ToString() != cellValue) {
                                            allMatch = false;
                                            break;
                                        }
                                    }
                                    if (allMatch) {
                                        exists = true;
                                        break;
                                    }
                                }

                                if (!exists) {
                                    DataRow newRow = dataTable.NewRow();
                                    foreach (var mapping in columnMapping) {
                                        var columnName = mapping.Value;
                                        var columnIndex = GetColumnIndex(worksheet, mapping.Key);
                                        var cellValue = row.Cell(columnIndex).Value.ToString();
                                        newRow[columnName] = cellValue;
                                    }
                                    dataTable.Rows.Add(newRow);
                                }
                            }

                            dataGridView1.DataSource = dataTable;
                        }
                        MessageBox.Show("Importação bem-sucedida!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Erro ao importar do Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private int GetColumnIndex(IXLWorksheet worksheet, string columnName) {
            var column = worksheet.Row(1).CellsUsed(c => c.Value.ToString() == columnName).FirstOrDefault();
            if (column != null) {
                return column.Address.ColumnNumber;
            }
            throw new Exception($"Coluna '{columnName}' não encontrada no Excel.");
        }

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e) {
        }
    }
}
