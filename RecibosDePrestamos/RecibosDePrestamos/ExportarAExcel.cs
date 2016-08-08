﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace Recibos
{
    public class ExportarAExcel
    {
        // Método que inicia el Excel
        public void startUp()
        {
            Microsoft.Office.Interop.Excel.Application appExcel =
                new Microsoft.Office.Interop.Excel.Application();   // Excel
            Workbook libro = appExcel.Workbooks.Add();   // Libro de Excel
            Worksheet hoja = libro.Worksheets.Add();   // nueva hoja de excel

            hoja.Name = "SEMANAL";
            ajustarTamanioFilas(hoja);
            ajustarAnchoEnFilas(hoja);
            ajustarBordes(hoja);
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