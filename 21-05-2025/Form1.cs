using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Google.Protobuf.WellKnownTypes;
using IronXL;

namespace _21_05_2025
{
    public partial class Form1 : Form
    {
        string ruta = AppDomain.CurrentDomain.BaseDirectory;
        int rowl = 0;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ruta += "file.xlsx";
        }

        private void btnConectar_Click(object sender, EventArgs e)
        {
            WorkBook workBook = WorkBook.Load(@ruta);
            WorkSheet workSheet = workBook.WorkSheets.First();

            int cellValue = workSheet["A2"].IntValue;
            foreach (var cell in workSheet["A1:C5"])
            {
                MessageBox.Show(cell.AddressString + " " + cell.Text);
            }
            decimal sum = workSheet["A2:A5"].Sum();
            MessageBox.Show(sum.ToString());

            decimal max = workSheet["A2:A5"].Max();
            MessageBox.Show(max.ToString());
            int i = 10;
            workSheet["A"+i.ToString()].Value = "8";

            workSheet.SetCellValue(1, 1, "prueba");
            workSheet.SaveAs(@ruta);
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            WorkBook workBook = WorkBook.Load(@ruta);
            WorkSheet workSheet = workBook.WorkSheets.First();
             if(txtMatricula.Text == "")
            {
                MessageBox.Show("Ingrese el valor de Matricula"); 
                txtMatricula.Focus();
            }
            else if (txtNombre.Text == "")
            {
                MessageBox.Show("Ingrese el valor de Nombre");
                txtNombre.Focus();
            }
            else if (txtCarrera.Text == "")
            {
                MessageBox.Show("Ingrese el valor de Carrera");
                txtCarrera.Focus();
            }
            else
            {
                workSheet.SetCellValue(rowl,1 , txtMatricula);
                workSheet.SetCellValue(rowl, 1, txtNombre);
                workSheet.SetCellValue(rowl, 1, txtCarrera);
                rowl++;
                txtMatricula.Clear();
                txtCarrera.Clear();
                txtNombre.Clear();
            }

                rowl++;
            workSheet.SaveAs(@ruta);
        }
    }
}
