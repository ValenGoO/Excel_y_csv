using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace prueba
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Buscar";
            openFileDialog1.Filter = "Archivos de imagen (*.jpg;*.jpeg;*.png;*.bmp)|*.jpg;*.jpeg;*.png;*.bmp|Todos los archivos (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string ruta = openFileDialog1.FileName;
                try
                {
                    // Cargar la imagen en el PictureBox
                    pictureBox1.Image = Image.FromFile(ruta);
                    pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage; // Ajustar el tamaño si es necesario
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al cargar la imagen: " + ex.Message);
                }
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            openFileDialog1.Title = "Buscar";
            openFileDialog1.Filter = "Archivos de Excel (*.xlsx)|*.xlsx|Archivos CSV (*.csv)|*.csv|Todos los archivos (*.*)|*.*";

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string ruta = openFileDialog1.FileName;
                try
                {
                    // Determinar la extensión del archivo
                    string extension = Path.GetExtension(ruta);

                    if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        // Cargar el archivo Excel
                        using (var package = new ExcelPackage(new FileInfo(ruta)))
                        {
                            var workbook = package.Workbook;
                            var worksheet = workbook.Worksheets[0]; // Leer la primera hoja
                            var dataTable = new DataTable();

                            // Asumir que la primera fila contiene los encabezados
                            int colCount = worksheet.Dimension.End.Column;
                            for (int i = 1; i <= colCount; i++)
                            {
                                dataTable.Columns.Add(worksheet.Cells[1, i].Text);
                            }

                            // Agregar las filas al DataTable
                            for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                            {
                                var newRow = dataTable.NewRow();
                                for (int col = 1; col <= colCount; col++)
                                {
                                    newRow[col - 1] = worksheet.Cells[row, col].Text;
                                }
                                dataTable.Rows.Add(newRow);
                            }

                            // Asignar el DataTable al DataGridView
                            dataGridView1.DataSource = dataTable;
                        }
                    }
                    else if (extension.Equals(".csv", StringComparison.OrdinalIgnoreCase))
                    {
                        // Leer el archivo CSV
                        var lines = File.ReadAllLines(ruta);
                        var dataTable = new DataTable();

                        // Asumir que la primera línea contiene los encabezados
                        string[] headers = lines[0].Split(',');
                        foreach (var header in headers)
                        {
                            dataTable.Columns.Add(header.Trim());
                        }

                        // Agregar las filas al DataTable
                        for (int i = 1; i < lines.Length; i++)
                        {
                            var row = lines[i].Split(',');
                            dataTable.Rows.Add(row);
                        }

                        // Asignar el DataTable al DataGridView
                        dataGridView1.DataSource = dataTable;
                    }
                    else
                    {
                        MessageBox.Show("Tipo de archivo no soportado.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error al cargar el archivo: " + ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DataTable dt = (DataTable)dataGridView1.DataSource;
            if (dt != null)
            {
                ExportarAExcel(dt);
            }
            else
            {
                MessageBox.Show("No hay datos para exportar.");
            }
        }
        private void ExportarAExcel(DataTable dataTable)
        {
            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Datos"); // Nombre de la hoja

                // Agregar encabezados
                for (int i = 0; i < dataTable.Columns.Count; i++)
                {
                    worksheet.Cells[1, i + 1].Value = dataTable.Columns[i].ColumnName; // Encabezados
                }

                // Agregar datos
                for (int i = 0; i < dataTable.Rows.Count; i++)
                {
                    for (int j = 0; j < dataTable.Columns.Count; j++)
                    {
                        worksheet.Cells[i + 2, j + 1].Value = dataTable.Rows[i][j]; // Datos
                    }
                }

                // Guardar archivo
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Archivos de Excel (*.xlsx)|*.xlsx",
                    Title = "Guardar archivo Excel"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    FileInfo file = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(file); // Guardar el paquete
                    MessageBox.Show("Archivo exportado exitosamente.");
                }
            }
        }
    }
}
