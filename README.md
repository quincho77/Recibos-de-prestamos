# Recibos-de-prestamos
En este repositorio se desarrollará una aplicación que envía información a una hoja de Excel, para posteriormente ser tratada desde allí o utilizarla como simple reporte. 

## Descripción
En este repositorio se desarrollará una aplicación que envía información a una hoja de Excel, para posteriormente ser tratada desde allí o utilizarla como simple reporte. La aplicación se realizará .NET utilizando dos tipos de proyectos .NET; -Aplicación de formularios de Windows, - Biblioteca de clases.
#### Herramientas
  Libreria Microsoft.Office.Iterop  14.0
###### Nota: Otra herramienta que funciona para los mismos fines es EPPlus

## Perspectiva del producto
Este es un producto independiente para automatizar los procesos que actualmente se realizan manualmente en hojas de calculo con la finalidad de permtir al negocio familiar(Prestamos de gabeta) gestionar los recibos que se le entregan a los clientes como comprobante de pago. Como anteriormete se señala, esto se realiza mediante hojas de cálculo en Excel, la idea es poder generar el Excel desde una vetana de Windows.
El documento se generará mediante un evento de usuario, el documento contiene una hoja donde están todos los recibos que se le entregarán como comprobante de pago al cliente, identificados por fecha; en total un prestamo tiene un plazo de 18 semanas para ser cancelado por ende son 18 recibos o tiquetes los que se generan, cada semana al momento del pago se le entregará el tiquete al cliente correspondiente a la semana

## Anexos
Asi tendría que lucir un documento final.
![imagen recibo](https://cloud.githubusercontent.com/assets/12851489/17451458/3f13ed4e-5b24-11e6-87c6-c07b011471e3.png)
