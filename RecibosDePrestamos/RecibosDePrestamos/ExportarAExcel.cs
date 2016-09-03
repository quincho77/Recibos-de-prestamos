using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

using Microsoft.Office.Interop.Excel;


namespace Recibos
{
    public class ExportarAExcel
    {
        Application appExcel;
        Workbook libro;
        Worksheet hoja;

        // Método que inicia el Excel
        public void startUp(Recibo nuevoRecibo)
        {
            appExcel = new Application();   // Excel
            libro = appExcel.Workbooks.Add();   // Libro de Excel
            hoja = libro.Worksheets.Add();   // nueva hoja de excel

            hoja.PageSetup.LeftMargin = 10.0;
            hoja.PageSetup.RightMargin = 0.25;
            hoja.Name = "SEMANAL";

            ajustarTamanioFilas(hoja);
            ajustarAnchoEnFilas(hoja);
            ajustarBordes(hoja);
            insertarNombre(hoja, nuevoRecibo.Nombre);
            insertarSemana(hoja, nuevoRecibo.Semana, nuevoRecibo.Fecha);
            insertarEtiquetas(hoja);
            insertarFecha(hoja, nuevoRecibo.Fecha);
            calcularSaldos(hoja, nuevoRecibo);
            insertarSaldoAnterior(hoja, nuevoRecibo);
            insertarAbono(hoja, nuevoRecibo);
            insertarSaldoActual(hoja, nuevoRecibo);

            appExcel.Visible = true;
        }// fin del método start_Up

        // El siguiente método cierra las instancias de Excel
        public void closeUp()
        {
            libro.Close(true, Type.Missing, Type.Missing);
            libro = null;
            appExcel.Quit();
            appExcel = null;
        }// fin del método closeUp

        // El siguiente método prepara el tamaño de las filas en la hoja de excel
        public void ajustarTamanioFilas(Worksheet hoja)
        {
            int cursor = 1;    // Nos indica en fila esta posicionado

            Range fila = hoja.Rows["1"];   // Aquí el indice de la fila equivale al cursor
            fila.RowHeight = 9;     // tamaño de fila de 9pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);

            fila = hoja.Rows[cursor];
            fila.RowHeight = 9.75;  // tamaño de fila de 9.75pts
            cursor++;

            cursor = ajustarTamanioFilasRecibo(cursor, hoja);
        }// fin del método ajustarFilas

        // El siguiente método ajusta el ancho de las filas en la columna A
        public void ajustarAnchoEnFilas(Worksheet hoja)
        {
            Range columna = hoja.Columns["A"];
            columna.ColumnWidth = 30; // columna A ancho de 30pts

            columna = hoja.Columns["B"];
            columna.ColumnWidth = 14;  // columna B ancho de 14pts

            columna = hoja.Columns["C"];
            columna.ColumnWidth = 0.92;  // columna C ancho de 0.92pts

            columna = hoja.Columns["D"];
            columna.ColumnWidth = 30;    // columna D ancho de 30pts

            columna = hoja.Columns["E"];
            columna.ColumnWidth = 14;  // columna E ancho de 14pts

            columna = hoja.Columns["F"];
            columna.ColumnWidth = 0.75;  // columna F ancho de 0.92pts
        }// fin del método ajustarAnchoEnFilas

        // El siguiente método coloca bordes alrededor de las celdas que conforman
        // el recibo
        public void ajustarBordes(Worksheet hoja)
        {
            hoja.get_Range("A1", "B55").Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            hoja.get_Range("A2", "C54").Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;
            hoja.get_Range("D1", "E55").Borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            hoja.get_Range("C2", "F54").Borders[XlBordersIndex.xlInsideVertical].LineStyle = XlLineStyle.xlContinuous;

            hoja.get_Range("A2", "C54").Borders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
        }// fin del método ajustarBordes

        // El siguiente método inserta el nombre del cliente en el Recibo
        public void insertarNombre(Worksheet hoja, String nombre)
        {
            int fila = 2;    // representa el posicionamiento de la fila en la hoja

            while (fila <= 50)
            {
                hoja.Cells[fila, "A"] = nombre.ToUpper();
                hoja.Cells[fila, "D"] = nombre.ToUpper();

                // con las siguiente lineas damos formato a las celdas
                hoja.get_Range("A" + fila, "D" + fila).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + fila, "D" + fila).Font.Name = "Arial";
                hoja.get_Range("A" + fila, "D" + fila).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                if (fila == 50)   // el último recibo
                    hoja.Range["D" + fila].Font.Italic = true;

                fila = fila + 6;  // incremento
            }// fin del while
        }// fin del método insertarNombre

        // El siguiente método inserta el numero de la semana del prestamo
        // en el recibo
        public void insertarSemana(Worksheet hoja, int semana, DateTime fecha)
        {
            int fila = 2;   // representa el posicionamiento de la fila en la hoja

            while (fila <= 50)
            {
                hoja.Cells[fila, "B"] = semana + " SEMANA";
                hoja.Cells[fila, "E"] = (semana + 9) + " SEMANA";

                // con las siguiente lineas damos formato a las celdas
                hoja.get_Range("E" + fila).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("E" + fila).Font.Name = "Arial";
                hoja.get_Range("E" + fila).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                if (fila == 50)
                {
                    hoja.Cells[fila, "E"] = calcularReciboParcial(
                        fecha.ToString("dddd")) + " PARCIALES";

                    // formto de celda
                    hoja.get_Range("E" + fila).Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    hoja.get_Range("E" + fila).Font.Bold = true;
                }// fin del if

                fila = fila + 6;
                semana++;
            }// fin del while

        }// fin del método insertarSemana

        // El siguiente método inserta las etiquetas Fecha, Saldo Anterior, Abono,
        // Saldo Actual en el Recibo
        public void insertarEtiquetas(Worksheet hoja)
        {
            int fila = 3; // representa el posicionamiento de la fila en la hoja

            while (fila <= 51)
            {
                hoja.Cells[fila, "A"] = "Fecha";   
                hoja.Cells[fila, "D"] = "Fecha";

                hoja.Cells[fila + 1, "A"] = "Saldo Anterior";
                hoja.Cells[fila + 1, "D"] = "Saldo Anterior";

                hoja.Cells[fila + 2, "A"] = "Abono";
                hoja.Cells[fila + 2, "D"] = "Abono";

                hoja.Cells[fila + 3, "A"] = "Saldo Actual";
                hoja.Cells[fila + 3, "D"] = "Saldo Actual";

                // con las siguiente lineas damos formato a las celdas
                hoja.get_Range("A" + fila, "E" + fila).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + fila, "E" + fila).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 1), "E" + (fila + 1)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 1), "E" + (fila + 1)).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 2), "E" + (fila + 2)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 2), "E" + (fila + 2)).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 3), "E" + (fila + 3)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 3), "E" + (fila + 3)).Font.Name = "Arial";

                fila = fila + 6;
            }// fin del while
        }// fin del método insertarEtiquetas

        // El siguiente método inserta las fechas en cada recibo
        private void insertarFecha(Worksheet hoja, DateTime fecha)
        {
            int fila = 3;   // representa el posicionamiento de la fila en la hoja
            int parcial = calcularReciboParcial(fecha.ToString("dddd"));

            if (parcial != 0)
            {
                hoja.Cells[51, "E"] = fecha.ToString("dd-MMM-yy");
                fecha = fecha.AddDays(parcial);
                hoja.get_Range("E51").HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            }// fin del if 

            while (fila <= 51)
            {
                hoja.Cells[fila, "B"] = fecha.ToString("dd-MMM-yy");
                hoja.get_Range("B" + fila).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                fecha = fecha.AddDays(7);
                fila = fila + 6;
            }// fin del while

            fila = 3;   // representa el posicionamiento de la fila en la hoja

            while (fila <= 45)
            {
                hoja.Cells[fila, "E"] = fecha.ToString("dd-MMM-yy");
                hoja.get_Range("E" + fila).HorizontalAlignment =
                    Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;

                fecha = fecha.AddDays(7);
                fila = fila + 6;
            }// fin del while
        }// fin del método insertarFecha

        // El siguiente método inserta el Saldo Anterior en la hoja
        private void insertarSaldoAnterior(Worksheet hoja, Recibo objRecibo)
        {
            int fila = 4;   // representa el posicionamiento de la fila en la hoja
            int parcial = calcularReciboParcial(objRecibo.Fecha.ToString("dddd"));
            int indice = 0;   // indece para recorrer arreglos

            if (parcial != 0)
            {
                hoja.Cells[52, "E"] = objRecibo.SaldoAnterior[indice];
                hoja.get_Range("E52").NumberFormat = "#,##0.00";
                indice++;
            }// fin del if

            while (fila <= 52)
            {
                hoja.Cells[fila, "B"] = objRecibo.SaldoAnterior[indice];
                hoja.get_Range("B" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while

            fila = 4;   // reposicion

            while (fila <= 46)
            {
                hoja.Cells[fila, "E"] = objRecibo.SaldoAnterior[indice];
                hoja.get_Range("E" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while

        }// fin del método insertarSaldoAnterior


        // El siguiente método inserta el Abono en la hoja
        private void insertarAbono(Worksheet hoja, Recibo objRecibo)
        {
            int fila = 5;   // representa el posicionamiento de la fila en la hoja
            int parcial = calcularReciboParcial(objRecibo.Fecha.ToString("dddd"));
            int indice = 0;   // indece para recorrer arreglos

            if (parcial != 0)
            {
                hoja.Cells[53, "E"] = objRecibo.Abono[indice];
                hoja.get_Range("E53").NumberFormat = "#,##0.00";
                indice++;
            }// fin del if

            while (fila <= 53)
            {
                hoja.Cells[fila, "B"] = objRecibo.Abono[indice];
                hoja.get_Range("B" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while

            fila = 5;   // reposicion

            while (fila <= 47)
            {
                hoja.Cells[fila, "E"] = objRecibo.Abono[indice];
                hoja.get_Range("E" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while
        }// fin del método insertarAbono

        // El siguiente método insertar el Saldo Actual en la hoja
        private void insertarSaldoActual(Worksheet hoja, Recibo objRecibo)
        {
            int fila = 6;   // representa el posicionamiento de la fila en la hoja
            int parcial = calcularReciboParcial(objRecibo.Fecha.ToString("dddd"));
            int indice = 0;   // indece para recorrer arreglos

            if (parcial != 0)
            {
                hoja.Cells[54, "E"] = objRecibo.SaldoActual[indice];
                hoja.get_Range("E54").NumberFormat = "#,##0.00";
                indice++;
            }// fin del if

            while (fila <= 54)
            {
                hoja.Cells[fila, "B"] = objRecibo.SaldoActual[indice];
                hoja.get_Range("B" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while

            fila = 6;   // reposicion

            while (fila <= 48)
            {
                hoja.Cells[fila, "E"] = objRecibo.SaldoActual[indice];
                hoja.get_Range("E" + fila).NumberFormat = "#,##0.00";
                fila = fila + 6;
                indice++;
            }// fin del while
        }// fin del método insertarSaldoActual

        // El siguiente método realiza los calculos del Saldo Anterior, Saldo Actual 
        // Abono, para posteriormente insertarlos en la hoja
        private void calcularSaldos(Worksheet hoja,  Recibo objRecibo)
        {
            objRecibo.SaldoAnterior[0] = (objRecibo.Monto * 0.2) + objRecibo.Monto;
            double cuotaDiaria = objRecibo.Monto / 100;
            int parcial = calcularReciboParcial(objRecibo.Fecha.ToString("dddd"));
            double cuotaSemanal = (objRecibo.Monto * 7) / 100;

            if (parcial != 0)
                objRecibo.Abono[0] = cuotaDiaria * parcial;

            else
                objRecibo.Abono[0] = cuotaSemanal;

            for (int i = 1; i < objRecibo.SaldoAnterior.Length; i++)
            {
                objRecibo.SaldoActual[i - 1] = objRecibo.SaldoAnterior[i - 1] - objRecibo.Abono[i - 1];
                objRecibo.SaldoAnterior[i] = objRecibo.SaldoActual[i - 1];

                if (objRecibo.SaldoAnterior[i] >= cuotaSemanal)
                {
                    if (objRecibo.SaldoAnterior[i] < (cuotaSemanal * 2) && parcial == 0)
                    {
                        objRecibo.Abono[i] = objRecibo.SaldoAnterior[i];
                        objRecibo.SaldoActual[i] = 0.0;
                    }// fin del if

                    else
                        objRecibo.Abono[i] = cuotaSemanal;
                }// fin del if

                else
                {
                    objRecibo.Abono[i] = objRecibo.SaldoAnterior[i];
                    objRecibo.SaldoActual[i] = 0.0;
                }// fin del else
            }// fin del for
        }// fin del método insertarSaldoAnterior

        // Existe un recibo que se toma en cuenta dependiendo del día de la 
        // semana, solo si el día de la semana es Martes el recibo no es tomado
        // de igual manera este recibo siempre lleva un formato especial.
        // El siguiente método crea y da formato a ese recibo.
        public int calcularReciboParcial(String fecha)
        {
            int parcial = 0;

            switch (fecha)
            {
                case "Monday":    // si el dia es lunes se rebajan 1 parcial(un día)
                    parcial = 1;        
                    break;

                case "Tuesday":  // si es miercoles se rebajan 7 parciles(7 días)
                    parcial = 0;
                    break;

                case "Wednesday":  // si es miercoles se rebajan 6 parciles(6 días)
                    parcial = 6;
                    break;

                case "Thursday":  // si es miercoles se rebajan 5 parciles(5 días)
                    parcial = 5;
                    break;

                case "Friday": // si es miercoles se rebajan 4 parciles(4 días)
                    parcial = 4;
                    break;

                case "Saturday": // si es miercoles se rebajan 3 parciles(3 días)
                    parcial = 3;
                    break;

                case "Sunday":  // si es miercoles se rebajan 2 parciles(2 días)
                    parcial = 2;
                    break;
            }// fin del switch

            return parcial;
        }// fin del método crearReciboParcial



        // DE ESTE BLOQUE DE CÓDIGO EN ADELANTE SE ESCRIBEN LOS MÉTODOS QUE SON
        // INVOCADOS POR EL MÉTODO ajustarFilas

        // El siguiente método ajusta el tamaño de las filas que conforman
        // parte del recibo
        public int ajustarTamanioFilasRecibo(int indice, Worksheet hoja)
        {
            for (int i = indice; i < indice + 5; i++)
            {
                Range fila = hoja.Rows[i];
                fila.RowHeight = 13.50;       // tamaño de fila de 13.50 pts
            }// fin del for

            indice = indice + 5;     // posicionamos el cursor
            return indice;
        }// fin del método ajustarTamanioFilasRecibo

    }// fin de la clase ExportarAExcel

}// fin del namespace Recibos
