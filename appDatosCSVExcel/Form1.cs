using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Text;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Office.Interop.Excel;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;

namespace appDatosCSVExcel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            lstvDatos.View = System.Windows.Forms.View.Details;
        }

        private void btnAbrir_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialogo = new OpenFileDialog();
            if (dialogo.ShowDialog() != DialogResult.OK)
            { return; }
            lstvDatos.Clear();
            string rutaArchivo = dialogo.FileName;

            StreamReader sr = new StreamReader(rutaArchivo, Encoding.GetEncoding(1252));
            string columnas = sr.ReadLine();
            string[] columna = columnas.Split('|');

            for (int i = 0; i < columna.Length; i++)
            { lstvDatos.Columns.Add(columna[i]); }

            string renglon;
            while ((renglon = sr.ReadLine()) != null)
            {
                string[] datos = renglon.Split('|');
                ListViewItem item = new ListViewItem(datos[0]);

                for (int i = 1; i < datos.Length; i++)
                { item.SubItems.Add(datos[1]); }

                lstvDatos.Items.Add(item);
            }

            sr.Close();
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Comma Separated Values | *.csv | Texto | *.txt | Excel | *.xlsx";

            if (sfd.ShowDialog() != DialogResult.OK)
            { return; }

            StreamWriter sw = new StreamWriter(sfd.FileName);

            foreach (ListViewItem item in lstvDatos.Items)
            {
                int i = 1;
                foreach (ListViewItem.ListViewSubItem subItem in item.SubItems)
                {
                    if (i == lstvDatos.Columns.Count)
                    {
                        sw.Write(subItem.Text);
                        sw.WriteLine();
                        i = 0;
                    }
                    else
                    { sw.Write(subItem.Text + ","); i++; }
                }
            }
            sw.Close();
            MessageBox.Show("Done");
        }

        private void btnGuardarExcel_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (sfd.ShowDialog() != DialogResult.OK)
            { return; }

            if (sfd.FileName != "")
            {
                Excel.Application excel = new Excel.Application();
                Excel.Workbook workbook = excel.Workbooks.Add(Type.Missing);
                Excel.Worksheet worksheet = null;

                try
                {
                    worksheet = workbook.ActiveSheet;
                    worksheet.Name = "ListView Data";

                    int row = 1;
                    int col = 1;

                    foreach (ColumnHeader column in lstvDatos.Columns)
                    {
                        worksheet.Cells[row, col] = column.Text;
                        col++;
                    }
                    row++;

                    foreach (ListViewItem item in lstvDatos.Items)
                    {
                        col = 1;
                        foreach (ListViewItem.ListViewSubItem subitem in item.SubItems)
                        {
                            worksheet.Cells[row, col] = subitem.Text;
                            col++;
                        }
                        row++;
                    }

                    workbook.SaveAs(sfd.FileName);
                    Process.Start(sfd.FileName);
                }

                catch (Exception ex)
                { MessageBox.Show("Error: " + ex.Message); }

                excel.Quit();
                workbook = null;
                excel = null;
            }
        }

        private void btnExcelOpenXML_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            if (sfd.ShowDialog() != DialogResult.OK)
            { return; }

            if (sfd.FileName != "")
            {
                SpreadsheetDocument document = SpreadsheetDocument.Create(sfd.FileName, SpreadsheetDocumentType.Workbook);
                // Agregar una hoja de trabajo
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new DocumentFormat.OpenXml.Spreadsheet.Worksheet(new SheetData());
                DocumentFormat.OpenXml.Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new DocumentFormat.OpenXml.Spreadsheet.Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "ListView" };
                sheets.Append(sheet);

                // Obtener la colección de celdas
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Recorrer las columnas y filas de la ListView
                for (int i = 0; i < lstvDatos.Columns.Count; i++)
                {
                    // Crear una fila para los encabezados
                    if (i == 0)
                    {
                        Row row = new Row() { RowIndex = 1 };
                        sheetData.Append(row);
                    }
                    // Obtener el encabezado de la columna
                    string headerText = lstvDatos.Columns[i].Text;
                    // Crear una celda para el encabezado
                    Cell headerCell = new Cell();
                    headerCell.CellReference = GetColumnName(i + 1) + "1";
                    headerCell.DataType = CellValues.String;
                    headerCell.CellValue = new CellValue(headerText);
                    
                    // Agregar la celda al final de la fila
                    Row headerRow = (Row)sheetData.ChildElements.GetItem(0);
                    headerRow.AppendChild(headerCell);

                    // Recorrer las filas de la columna
                    for (int j = 0; j < lstvDatos.Items.Count; j++)
                    {
                        // Crear una fila para los datos
                        if (i == 0)
                        {
                            Row row = new Row() { RowIndex = (uint)(j + 2) };
                            sheetData.Append(row);
                        }
                        // Obtener el valor del dato
                        string dataText = lstvDatos.Items[j].SubItems[i].Text;
                        // Crear una celda para el dato
                        Cell dataCell = new Cell();
                        dataCell.CellReference = GetColumnName(i + 1) + (j + 2);
                        dataCell.DataType = CellValues.String;
                        dataCell.CellValue = new CellValue(dataText);

                        // Agregar la celda al final de la fila
                        Row dataRow = (Row)sheetData.ChildElements.GetItem(j + 1);
                        dataRow.AppendChild(dataCell);
                    }
                }
                // Guardar y cerrar el documento
                workbookPart.Workbook.Save();
                document.Close();
                Process.Start(sfd.FileName);
            }
        }
        // Método auxiliar para obtener el nombre de la columna según su índice
        private string GetColumnName(int index)
        {
            int dividend = index;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }
    }
}

