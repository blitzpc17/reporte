using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace reporte
{
    public partial class frmReporte : Form
    {
        Excel.Application ReporteApp = default(Excel.Application);
        Excel.Workbook Libro = default(Excel.Workbook);
        Excel.Worksheet HojaReporte = default(Excel.Worksheet);
        private string ruta = "";


        public frmReporte()
        {
            InitializeComponent();
        }

        private void btnLeer_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog(this);
            ruta = openFileDialog1.FileName;
            btnCargarArchivo.Enabled = false;
            backgroundWorker1.RunWorkerAsync();
            //LeerTxt(ruta);
        }

        private void LeerTxt(string ruta)
        {
            ReporteApp = new Excel.Application();
            ReporteApp.Visible = true;

            Libro = ReporteApp.Workbooks.Add();
            HojaReporte = Libro.Worksheets[1];
            HojaReporte.Visible = Excel.XlSheetVisibility.xlSheetVisible;
            HojaReporte.Activate();


            int numeroRows = System.IO.File.ReadLines(ruta).Count();
            int countRows = 1;
            foreach ( string linea in System.IO.File.ReadLines(ruta))
            {
                string ln = linea;
                if (countRows == 1)
                {
                    HojaReporte.Cells[countRows, "A"] = "Cuenta";
                    HojaReporte.Cells[countRows, "B"] = "Sub Cta";
                    HojaReporte.Cells[countRows, "C"] = "Concepto";
                    HojaReporte.Cells[countRows, "D"] = "Saldo según balanza";
                    HojaReporte.Cells[countRows, "E"] = "Saldos s/Audit.";
                    HojaReporte.Cells[countRows, "F"] = "Diferencia";
                    HojaReporte.Cells[countRows, "G"] = "Marca Auditoria";
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
                                    HojaReporte.Cells[countRows - 1, "C"] = descrpcion;
                                }
                            }

                        }



                        /*
                        decimal salida = 0;

                        if (i == 0)
                        {
                            HojaReporte.Cells[countRows-1, "A"] = renglon[i].ToString();
                        }
                        else if (!string.IsNullOrWhiteSpace(renglon[i]) && !decimal.TryParse(renglon[i], out salida))
                        {
                            descrpcion += Convert.ToString(renglon[i]) + " ";
                            HojaReporte.Cells[countRows-1, "B"] = descrpcion;
                        }
                        else if (decimal.TryParse(renglon[i], out salida))
                        {
                            columnasImportes++;
                            if (columnasImportes > 4) break;

                            switch (columnasImportes)
                            {
                                case 1:
                                    HojaReporte.Cells[countRows-1, "C"] = salida;
                                    break;
                                case 2:
                                    HojaReporte.Cells[countRows-1, "D"] = salida;
                                    break;
                                case 3:
                                    HojaReporte.Cells[countRows-1, "E"] = salida;
                                    break;
                                case 4:
                                    HojaReporte.Cells[countRows-1, "F"] = salida;
                                    break;
                            }
                        */



                    }
                }
               
                countRows++;
               
            


                Console.WriteLine(linea);
            }
        }

        private void btnSalir_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            LeerTxt(ruta);
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("Reporte generado correctamente.", "Aviso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            btnCargarArchivo.Enabled = true;
            this.Focus();
        }
    }
}
