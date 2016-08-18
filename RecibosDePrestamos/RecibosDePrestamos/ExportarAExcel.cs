using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Recibos
{
    public class ExportarAExcel
    {
        // Método que inicia el Excel
        public void startUp(Recibo nuevoRecibo)
        {
            Microsoft.Office.Interop.Excel.Application appExcel =
                new Microsoft.Office.Interop.Excel.Application();   // Excel
            Workbook libro = appExcel.Workbooks.Add();   // Libro de Excel
            Worksheet hoja = libro.Worksheets.Add();   // nueva hoja de excel

            hoja.Name = "SEMANAL";
            ajustarTamanioFilas(hoja);
            ajustarAnchoEnFilas(hoja);
            ajustarBordes(hoja);
            insertarNombre(hoja, nuevoRecibo.Nombre);
            insertarSemana(hoja, nuevoRecibo.Semana, nuevoRecibo.Fecha);
            insertarEtiquetas(hoja);
            
            appExcel.Visible = true;
        }// fin del método start_Up

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
                hoja.get_Range("A" + fila, "D" + fila).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + fila, "D" + fila).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 1), "D" + (fila + 1)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 1), "D" + (fila + 1)).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 2), "D" + (fila + 2)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 2), "D" + (fila + 2)).Font.Name = "Arial";

                hoja.get_Range("A" + (fila + 3), "D" + (fila + 3)).Font.Size = 8;   // tamaño de letra de 8pt
                hoja.get_Range("A" + (fila + 3), "D" + (fila + 3)).Font.Name = "Arial";

                fila = fila + 6;
            }// fin del while
        }// fin del método insertarEtiquetas



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
                    parcial = 7;
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
