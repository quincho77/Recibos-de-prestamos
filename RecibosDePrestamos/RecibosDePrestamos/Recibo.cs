using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Recibos
{
    public class Recibo
    {
        private String nombre;
        private double monto;
        private String caja;
        private DateTime fecha;
        private double saldoAnterior;
        private double saldoActual;
        private double abono;
        private int semana;


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

        public DateTime Fecha
        {
            set { fecha = value; }
            get { return fecha; }
        }// fin de la propiedad fecha

        public double SaldoAnterior
        {
            set { saldoAnterior = value; }
            get { return saldoAnterior; }
        }// fin de la propiedad saldoAnterior

        public double SaldoActual
        {
            set { saldoActual = value; }
            get { return saldoActual; }
        }// fin de la propiedad SaldoActual

        public double Abono
        {
            set { abono = value; }
            get { return abono; }
        }// fin de la propiedad abono

        public int Semana
        {
            set { semana = value; }
            get { return semana;  }
        }// fin de la propiedad fecha

    }// fin de la clase Recibo

}// fin del namespace Recibos
