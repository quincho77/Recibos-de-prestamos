using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Recibos;

namespace GUI
{
    public partial class frmRecibos : Form
    {
        public frmRecibos()
        {
            InitializeComponent();
        }

        private void btnVerExcel_Click(object sender, EventArgs e)
        {
            Recibo objRecibo = obtenerDatosRecibo();

            ExportarAExcel objeto = new ExportarAExcel();
            objeto.startUp(objRecibo);
            //MessageBox.Show(objRecibo.Fecha.ToString("dddd"));
        }// fin del metodo btnVerExcel

        private Recibo obtenerDatosRecibo()
        {
            Recibo objRecibo = new Recibo();
            objRecibo.Nombre = txtNombre.Text;
            objRecibo.Monto = Convert.ToDouble(txtMonto.Text);
            objRecibo.Caja = txtCaja.Text;
            objRecibo.Fecha = DateTime.Today;


            return objRecibo;
        }// fin del método obtenerDatosRecibo
    }
}
