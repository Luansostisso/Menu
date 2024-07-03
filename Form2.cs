using System;
using System.Data;
using System.Windows.Forms;
using FirebirdSql.Data.FirebirdClient;
using ClosedXML.Excel;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Linq;
using Fiscal.Classe;

namespace Desafio2 {
    public partial class Form2 : Form {
        public Form2() {
            InitializeComponent();
            ContextMenuStrip contextMenu = new ContextMenuStrip();

            ToolStripMenuItem exportToExcel = new ToolStripMenuItem("Extrair para Excel");
            exportToExcel.Click += ExportToExcel_Click;
            contextMenu.Items.Add(exportToExcel);

            ToolStripMenuItem importFromExcel = new ToolStripMenuItem("Importar .xls");
            importFromExcel.Click += ImportFromExcel_Click;
            contextMenu.Items.Add(importFromExcel);

            dataGridView1.ContextMenuStrip = contextMenu;

            LoadDataAsync();
        }

        private async Task LoadDataAsync() {
            try {
                using (var context = new DataContext.Contexto()) {
                    await context.Database.OpenConnectionAsync();

                    string selectSql = "SELECT * FROM TCliente";

                    using (FbCommand cmd = new FbCommand(selectSql, (FbConnection)context.Database.GetDbConnection())) {
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

        private async void ExportToExcel_Click(object sender, EventArgs e) {
            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx" }) {
                if (sfd.ShowDialog() == DialogResult.OK) {
                    try {
                        await Task.Run(() => ExportToExcel(sfd.FileName));
                        MessageBox.Show("Exportação bem-sucedida!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Erro ao exportar para Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void ExportToExcel(string fileName) {
            using (var context = new DataContext.Contexto()) {
                string selectSql = "SELECT * FROM TCliente";

                using (FbCommand cmd = new FbCommand(selectSql, (FbConnection)context.Database.GetDbConnection())) {
                    using (FbDataReader reader = cmd.ExecuteReader()) {
                        using (var workbook = new XLWorkbook()) {
                            var worksheet = workbook.Worksheets.Add("Sheet1");

                            for (int i = 0; i < reader.FieldCount; i++) {
                                worksheet.Cell(1, i + 1).Value = reader.GetName(i);
                            }

                            int row = 2;
                            while (reader.Read()) {
                                for (int i = 0; i < reader.FieldCount; i++) {
                                    worksheet.Cell(row, i + 1).Value = reader.GetValue(i)?.ToString();
                                }
                                row++;
                            }

                            workbook.SaveAs(fileName);
                        }
                    }
                }
            }
        }

        private async void ImportFromExcel_Click(object sender, EventArgs e) {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx" }) {
                if (ofd.ShowDialog() == DialogResult.OK) {
                    try {
                        await Task.Run(() => ImportFromExcelAsync(ofd.FileName));
                        await LoadDataAsync();
                        MessageBox.Show("Importação bem-sucedida!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (DbUpdateException ex) {
                        MessageBox.Show($"Erro ao importar do Excel: {ex.InnerException?.Message ?? ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    catch (Exception ex) {
                        MessageBox.Show($"Erro ao importar do Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private async Task ImportFromExcelAsync(string fileName) {
            using (var workbook = new XLWorkbook(fileName)) {
                var worksheet = workbook.Worksheet(1);
                var dataTable = dataGridView1.DataSource as DataTable ?? new DataTable();

                var columnMapping = new Dictionary<string, string>
                {
                    { "CPF", nameof(Cliente.CPF) },
                    { "CLIENTE", nameof(Cliente.TipoCliente) },
                    { "FANTASIA", nameof(Cliente.Fantasia) },
                    { "RG/IE", nameof(Cliente.RG) },
                    { "ENDEREÇO", nameof(Cliente.Endereco) },
                    { "COMPLEMENTO", nameof(Cliente.Complemento) },
                    { "BAIRRO", nameof(Cliente.Bairro) },
                    { "CEP", nameof(Cliente.CEP) },
                    { "NÚMERO", nameof(Cliente.Numero) },
                    { "MUNICÍPIO", nameof(Cliente.Cidade) },
                    { "UF", nameof(Cliente.UF) },
                    { "TELEFONE/CELULAR", nameof(Cliente.Telefone) },
                    { "E-MAIL", nameof(Cliente.Email) },
                    { "CONTROLE", nameof(Cliente.Controle) }
                };
               
                foreach (var mapping in columnMapping) {
                    if (!dataTable.Columns.Contains(mapping.Value)) {
                        dataTable.Columns.Add(mapping.Value);
                    }
                }
                using (var context = new DataContext.Contexto()) {
                    using (var transaction = context.Database.BeginTransaction()) {
                        try {
                            int maxControle = await context.cliente.MaxAsync(c => (int?)c.Controle) ?? 0;
                            
                            var existingClients = await context.cliente.ToListAsync();
                            var existingClientCPFs = new HashSet<string>(existingClients.Select(c => c.CPF));

                            foreach (var row in worksheet.RowsUsed()) {
                                if (row.RowNumber() == 1)
                                    continue;

                                // Verificar se já existe um cliente com o mesmo CPF
                                var cpfCell = row.Cell(GetColumnIndex(worksheet, "CPF"));
                                if (cpfCell == null || cpfCell.IsEmpty()) {
                                    continue;
                                }

                                var cpf = cpfCell.GetString();
                                if (existingClientCPFs.Contains(cpf)) {
                                    continue;
                                }
                                
                                var cliente = new Cliente();

                                foreach (var mapping in columnMapping) {
                                    var columnIndex = GetColumnIndex(worksheet, mapping.Key);
                                    var cell = row.Cell(columnIndex);

                                    if (cell == null || cell.IsEmpty()) {                                       
                                        continue;
                                    }

                                    var cellValue = cell.GetString();
                                    var property = cliente.GetType().GetProperty(mapping.Value);

                                    if (property != null && property.CanWrite) {
                                        property.SetValue(cliente, cellValue);
                                    }
                                }

                                // Preencher DataHoraCadastro com a data e hora atuais
                                cliente.DataHoraCadastro = DateTime.Now;

                                if (string.IsNullOrEmpty(cliente.Ativo)) {
                                    cliente.Ativo = "S"; 
                                }

                                
                                cliente.Controle = ++maxControle; 

                                context.cliente.Add(cliente);
                                existingClientCPFs.Add(cpf);
                            }

                            await context.SaveChangesAsync();
                            transaction.Commit(); 

                            await LoadDataAsync(); 
                            MessageBox.Show("Importação bem-sucedida!", "Informação", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex) {
                            transaction.Rollback(); // Reverter a transação em caso de erro
                            MessageBox.Show($"Erro ao importar do Excel: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }       
        private int GetColumnIndex(IXLWorksheet worksheet, string columnName) {
            var headerRow = worksheet.Row(1);
            foreach (var cell in headerRow.Cells()) {
                if (cell.GetString().Equals(columnName, StringComparison.OrdinalIgnoreCase)) {
                    return cell.Address.ColumnNumber;
                }
            }
            return -1; 
        }

    }
}