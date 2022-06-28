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
    public partial class Form1 : Form
    {
        Excel.Application ReporteApp = default(Excel.Application);
        Excel.Workbook Libro = default(Excel.Workbook);
        Excel.Worksheet HojaReporte = default(Excel.Worksheet);


        public Form1()
        {
            InitializeComponent();
        }

        private void btnLeer_Click(object sender, EventArgs e)
        {
            openFileDialog1 = new OpenFileDialog();
            openFileDialog1.ShowDialog(this);
            string ruta = openFileDialog1.FileName;
            LeerTxt(ruta);
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
            int countRows = 0;
            foreach ( string linea in System.IO.File.ReadLines(ruta))
            {
                string ln = linea;
                if (countRows == 0)
                {
                    HojaReporte.Cells[countRows + 1, "A"] = "Cuenta";
                    HojaReporte.Cells[countRows + 1, "B"] = "Descripción";
                    HojaReporte.Cells[countRows + 1, "C"] = "Saldo Ant.";
                    HojaReporte.Cells[countRows + 1, "D"] = "Cargos";
                    HojaReporte.Cells[countRows + 1, "E"] = "Abonos";
                    HojaReporte.Cells[countRows + 1, "F"] = "Saldo";
                }
                else if (countRows >= 2)
                {
                    String[] renglon = ln.ToString().Split(' ');
                    string descrpcion = "";
                    int columnasImportes = 0;
                    for (int i = 0; i < renglon.Length; i++)
                    {
                        decimal salida = 0;

                        if (i == 0)
                        {
                            HojaReporte.Cells[countRows, "A"] = renglon[i].ToString();
                        }
                        else if (!string.IsNullOrWhiteSpace(renglon[i]) && !decimal.TryParse(renglon[i], out salida))
                        {
                            descrpcion += Convert.ToString(renglon[i]) + " ";
                            HojaReporte.Cells[countRows, "B"] = descrpcion;
                        }
                        else if (decimal.TryParse(renglon[i], out salida))
                        {
                            columnasImportes++;
                            if (columnasImportes > 4) break;

                            switch (columnasImportes)
                            {
                                case 1:
                                    HojaReporte.Cells[countRows, "C"] = salida;
                                    break;
                                case 2:
                                    HojaReporte.Cells[countRows, "D"] = salida;
                                    break;
                                case 3:
                                    HojaReporte.Cells[countRows, "E"] = salida;
                                    break;
                                case 4:
                                    HojaReporte.Cells[countRows, "F"] = salida;
                                    break;
                            }

                        }
                    }
                }
                else if (countRows + 3 == numeroRows ) break;
                countRows++;
               
            


                Console.WriteLine(linea);
            }
        }
    }
}
