using System;
using System.Data;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using ClosedXML.Excel;
using Fiscal.Classe;
using System.ComponentModel;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

namespace Desafio2 {
    public partial class Form3 : Form {
        public Form3() {
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

                    string selectSql = "SELECT * FROM TEstoque";

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
                            DataTable dataTable = new DataTable();

                            bool firstRow = true;
                            foreach (var row in worksheet.Rows()) {
                                if (firstRow) {
                                    foreach (var cell in row.Cells()) {
                                        dataTable.Columns.Add(cell.Value.ToString());
                                    }
                                    firstRow = false;
                                }
                                else {
                                    dataTable.Rows.Add();
                                    int i = 0;
                                    foreach (var cell in row.Cells()) {
                                        dataTable.Rows[dataTable.Rows.Count - 1][i] = cell.Value.ToString();
                                        i++;
                                    }
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

        private void contextMenuStrip1_Opening(object sender, CancelEventArgs e) {
        }
    }
}
