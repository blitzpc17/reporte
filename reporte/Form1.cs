using System;
using System.ComponentModel;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;
using Workbook = Microsoft.Office.Interop.Excel.Workbook;
using Worksheet = Microsoft.Office.Interop.Excel.Worksheet;
using Borders = Microsoft.Office.Interop.Excel.Borders;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using SpreadsheetLight;

namespace reporte
{
    public partial class frmReporte : Form
    {
        Application ReporteApp;
        Workbook Libro;
        Worksheet HojaReporte;
        private string ruta = "", rutaExcel ="";
        private int countRows = 0;
        bool excelCargado = false;
        bool archivoCargado = false;


        public frmReporte()
        {
            InitializeComponent();
        }

        private void btnLeer_Click(object sender, EventArgs e)
        {
            ImportarArchivoCuentas();
        }

        private void LeerTxt(string ruta, string rutaExcel)
        {
            ReporteApp = new Application();
            Libro = ReporteApp.Workbooks.Open(rutaExcel);
            HojaReporte = Libro.Worksheets[0];
            
            countRows = 0;


            int totalRows = HojaReporte.Rows.Count;
            int totalCols = HojaReporte.Columns.Count;

            countRows = totalRows + 5;

            int numeroRows = System.IO.File.ReadLines(ruta).Count();
            
            foreach ( string linea in System.IO.File.ReadLines(ruta))
            {
                string ln = linea;
                if (countRows == 1)
                {
                    HojaReporte.Cells[countRows, "A"] = "CUENTA";
                    HojaReporte.Cells[countRows, "B"] = "SUBCTA";
                    HojaReporte.Cells[countRows, "C"] = "CONCEPTO";
                    HojaReporte.Cells[countRows, "D"] = "SALDO SEGÚN BALANZA";
                    HojaReporte.Cells[countRows, "E"] = "SALDO S/ AUDITORIA";
                    HojaReporte.Cells[countRows, "F"] = "DIFERENCIA";
                    HojaReporte.Cells[countRows, "G"] = "MARCA DE AUDITORIA";
                }
                else if (countRows == numeroRows - 4) break;
                else if (countRows > 2)
                {
                    String[] renglon = ln.ToString().Split(' ');
                    string descrpcion = "";
                    int columnasImportes = 0;
                    decimal salida = 0;
                    for (int i = 0; i < renglon.Length; i++)
                    {
                        if (i == 0 && decimal.TryParse(renglon[i].ToString(), out salida))
                        {
                            columnasImportes = 0;
                            HojaReporte.Cells[countRows-1, "A"] = renglon[i].ToString();
                        }
                        else if(i == 0 )
                        {
                            HojaReporte.Cells[countRows - 1, "B"] = renglon[i].ToString();
                        }else
                        {
                            if(decimal.TryParse(renglon[i].ToString(), out salida))
                            {
                                columnasImportes++;
                                if (columnasImportes == 4)
                                {
                                    HojaReporte.Cells[countRows - 1, "D"] = renglon[i].ToString();
                                    HojaReporte.Cells[countRows - 1, "E"] = renglon[i].ToString();
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrWhiteSpace(renglon[i].ToString()))
                                {
                                    descrpcion += Convert.ToString(renglon[i]) + " ";
                                    HojaReporte.Cells[countRows - 1, "C"] = (descrpcion);
                                }
                            }

                        }
                    }
                }
               
                countRows++;
               
            


                Console.WriteLine(linea);
            }
            PintarBordes();
        }


        private void PintarBordes()
        {

            Range rango = HojaReporte.get_Range("A1", "G" + (countRows-2));

            Range rangoDescripcion = HojaReporte.get_Range("C1", "C" + (countRows - 2));
            rangoDescripcion.Columns.AutoFit();
            rangoDescripcion.Font.Bold = true;

            Range rangoEncabezados = HojaReporte.get_Range("A1", "F1");
            rangoEncabezados.Font.Bold = true;
            rangoEncabezados.WrapText = true;

            HojaReporte.Cells[1, 7].WrapText = true;


            Borders border = rango.Borders;
            rango.Font.Size = 10;
            border[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlDot;
            border[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            border[XlBordersIndex.xlEdgeLeft].Weight = XlBorderWeight.xlThick;
            border[XlBordersIndex.xlEdgeRight].Weight = XlBorderWeight.xlThick;
            border[XlBordersIndex.xlEdgeTop].Weight = XlBorderWeight.xlThick;
            border[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;

        }
        private void btnSalir_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            LeerTxt(ruta, rutaExcel);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Reporte generado correctamente.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnCargarArchivo.Enabled = true;
            ReporteApp.Visible = true;
            this.Focus();
        }

        private void ImportarArchivoCuentas()
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog(this);
            ruta = openFileDialog1.FileName;
            if (string.IsNullOrEmpty(ruta))
            {
                MessageBox.Show("No ha cargado ningun archivo. Intentelo de nuevo", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

        }

        private void buttonCargarExcel_Click(object sender, EventArgs e)
        {
            string error ="";
            if (string.IsNullOrEmpty(txtHoja.Text)){
                error = "la posición de la Hoja en la cual se realizará la generacion de las cuentas contables.";
            }else if(string.IsNullOrEmpty(txtFilas.Text)) 
            {
                error = "el numero de la fila en la cual se iniciara a generar las cuentas contables.";
            }else if (string.IsNullOrEmpty(txtColumna.Text))
            {
                error = "la columna en donde se situara la generación de las cuentas contables.";
            }

            if (!string.IsNullOrEmpty(error))
            {
                MessageBox.Show("Falto ingresar " + error, "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
            }
            ImportarExcel(Convert.ToInt32(txtHoja.Text), Convert.ToInt32(txtFilas.Text));
        }

        private void ImportarExcel(int hoja, int filas)
        {
            if (string.IsNullOrEmpty(ruta))
            {
                MessageBox.Show("No ha cargado el Archivo de Cuentas. Cargue el archivo e intentelo de nuevo", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog(this);
            rutaExcel = openFileDialog1.FileName;
            if (string.IsNullOrEmpty(rutaExcel))
            {
                MessageBox.Show("No ha cargado ningun archivo. Intentelo de nuevo", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            excelCargado = true;
            SLDocument sl = new SLDocument(rutaExcel);
            var lstHojas = sl.GetSheetNames();
            Console.WriteLine(lstHojas[(hoja - 1)]);
            string NombreHoja = /*SLDocument.DefaultFirstSheetName;*/lstHojas[(hoja - 1)];
            sl.SelectWorksheet(NombreHoja);
            Console.WriteLine(NombreHoja);
            Console.WriteLine( sl.GetCurrentWorksheetName());  

            countRows = filas;
            sl.InsertRow(filas, 1);

            sl.SetCellValue("A"+countRows, "CUENTA");
            sl.SetCellValue("B"+countRows, "SUBCTA");
            sl.SetCellValue("C"+countRows, "CONCEPTO");
            sl.SetCellValue("D"+countRows, "SALDO SEGÚN BALANZA");
            sl.SetCellValue("E"+countRows, "SALDO S/ AUDITORIA");
            sl.SetCellValue("F"+countRows, "DIFERENCIA");
            sl.SetCellValue("G"+countRows, "MARCA DE AUDITORIA");

            
            

            SLStyle styleBordeGrueso = sl.CreateStyle();
            SLStyle styleBordeDoteado = sl.CreateStyle();
            SLStyle styleNegrita = sl.CreateStyle();
            SLStyle styleThin = sl.CreateStyle();
            SLStyle styleWrapText = sl.CreateStyle();

            styleBordeGrueso.Border.LeftBorder.BorderStyle = BorderStyleValues.Medium;
            styleBordeGrueso.Border.RightBorder.BorderStyle = BorderStyleValues.Medium;
            styleThin.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            styleBordeDoteado.Border.BottomBorder.BorderStyle = BorderStyleValues.Dotted;
            styleWrapText.SetWrapText(true);
            styleWrapText.Alignment.Horizontal= HorizontalAlignmentValues.Center;
            styleNegrita.SetFontBold(true);
            sl.SetCellStyle(countRows, 1, countRows, 7, styleNegrita);
            sl.SetCellStyle(countRows, 1, countRows, 7, styleWrapText);

            int numeroRows = System.IO.File.ReadLines(ruta).Count() - 2;
            sl.InsertRow(countRows+1, numeroRows);
            int contadorLineas = countRows;
            var txtCuentasContables = System.IO.File.ReadLines(ruta);
            int posLineaCuentaContable = 0;
            foreach (string linea in txtCuentasContables)
            {
                posLineaCuentaContable++;
                if (posLineaCuentaContable <= 2)
                {
                    countRows++;
                    continue;
                }
               // contadorLineas++;
                
                string ln = linea;
                if (contadorLineas == numeroRows - 2) break;
                else if (contadorLineas > 2)
                {
                    String[] renglon = ln.ToString().Split(' ');
                    string descrpcion = "";
                    int columnasImportes = 0;
                    decimal salida = 0;
                    for (int i = 0; i < renglon.Length; i++)
                    {
                        if (i == 0 && decimal.TryParse(renglon[i].ToString(), out salida))
                        {
                            columnasImportes = 0;
                            var contenido = renglon[i].ToString();
                            sl.SetCellValue("A" + (countRows-1), renglon[i].ToString());
                        }
                        else if (i == 0)
                        {
                            var contenido = renglon[i].ToString();
                            sl.SetCellValue("B" + (countRows-1), renglon[i].ToString());
                        }
                        else
                        {
                            if (decimal.TryParse(renglon[i].ToString(), out salida))
                            {
                                columnasImportes++;
                                if (columnasImportes == 4)
                                {
                                    var contenido = renglon[i].ToString();
                                    sl.SetCellValue("D" + (countRows-1), renglon[i].ToString());
                                    sl.SetCellValue("E" + (countRows-1), renglon[i].ToString());
                                }
                            }
                            else
                            {
                                if (!string.IsNullOrWhiteSpace(renglon[i].ToString()))
                                {
                                    var contenido = renglon[i].ToString();
                                    descrpcion += Convert.ToString(renglon[i]) + " ";
                                    descrpcion = descrpcion.Contains('\a') ? descrpcion.Replace('\a', ' '): descrpcion;
                                    sl.SetCellValue("C" + (countRows-1), (descrpcion));
                                }
                            }

                        }
                    }

                    sl.SetCellStyle(countRows - 1, 1, countRows - 1, 7, styleBordeGrueso);
                    sl.SetCellStyle(countRows - 1, 1, countRows - 1, 7, styleBordeDoteado);
                    sl.SetCellStyle(countRows - 1, 1, countRows - 1, 3, styleNegrita);
                }

                Console.WriteLine(linea);
                countRows++;
            }







            saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.ShowDialog();

            sl.SaveAs(saveFileDialog1.FileName/*+".xlsx"*/);



        }
    }
}
