using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RecibosDePrestamos
{
    class Recibo
    {
        private String nombre;
        private double monto;
        private String caja;
        private DateTime fecha;

        public Recibo()
        {
            
        }// fin del constructor de Recibo

        public String Nombre
        {
            set { nombre = value; }
            get { return nombre; }
        }// fin de la propiedad nombre

        public double Monto
        {
            set { monto = value; }
            get { return monto; }
        }// fin de la propiedad monto

        public String Caja
        {
            set { caja = value; }
            get { return caja; }
        }// fin de la propiedad caja

        private DateTime Fecha
        {
            set { fecha = value; }
            get { return fecha; }
        }// fin de la propiedad fecha

    }// fin de la clase Recibo

}// fin del namespace RecibosDePrestamos
