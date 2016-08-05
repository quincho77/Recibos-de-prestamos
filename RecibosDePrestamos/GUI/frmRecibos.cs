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
            ExportarAExcel objeto = new ExportarAExcel();
            objeto.startUp();
        }// fin del metodo btnVerExcel

    }
}
