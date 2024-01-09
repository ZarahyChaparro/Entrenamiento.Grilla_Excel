using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using objExcel = Microsoft.Office.Interop.Excel;

namespace Proyecto_Grilla
{
    public partial class Form1 : Form
    {
        string ruta = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);


        public Form1()
        {
            InitializeComponent();
        }

        private void btnExcel_Click(object sender, EventArgs e)
        {
            objExcel.Application objAplicacion = new objExcel.Application();
            Workbook objLibro = objAplicacion.Workbooks.Add(XlSheetType.xlWorksheet);

            Worksheet objHoja = (Worksheet)objAplicacion.ActiveSheet;

            objAplicacion.Visible = false;

            foreach (DataGridViewColumn columna in dataGridView1.Columns)
            {
                objHoja.Cells[1, columna.Index + 1] = columna.HeaderText;
                foreach (DataGridViewRow fila in dataGridView1.Rows)
                {
                    objHoja.Cells[fila.Index + 2, columna.Index + 1] = fila.Cells[columna.Index].Value;
                }
            }

            objLibro.SaveAs(ruta + "\\ProyectoEntrenamiento_Grilla.xlsx");
            objLibro.Close();
            objAplicacion.Quit();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.Rows.Add("Carmenza", "Chavez", "3223748594", "carmenza@gmail.com");
            dataGridView1.Rows.Add("Pablo", "Escoral", "3235468958", "escoral@gmail.com");

        }
    }
}
