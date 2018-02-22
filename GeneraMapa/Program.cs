using System;
using System.Collections.Generic;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace GeneraMapa
{
    class Program
    {
        // La funcionalidad de copiar contenido al portapapeles no es compatible por defecto en aplicaciones de consola.
        // Una vez creada la referencia, al ejecutarse aparecerá una excepción.
        // Para evitar esa excepción, se utiliza este comando, y así se puede ocupar funcionalidades que en un principio no están
        // disponibles para app de consola.
        [STAThread]

        static void Main(string[] args)
        {
            Console.WriteLine("Inicio proceso.");

            #region Variables
            // Variables:
            // Nombre Archivo Excel: Nombre del archivo EXCEL con la data debe estar ubicado en la misma ruta del exe.
            String Archivo_Excel = Path.Combine(Directory.GetCurrentDirectory(), "Catastro GEC.xlsx");

            // Nombre Archivo a generar: Nombre del archivo temporal que se usará como nexo.
            String Archivo_Temp = Path.Combine(Directory.GetCurrentDirectory(), "SalidaAuxiliar.txt");

            // Desde Columna Excel: Número de columna desde el que se comenzará a tomar los datos. (Las primeras columnas no son parte del FreeMind)
            int Excel_Columna_Desde = 3;  // "Menu"

            // Desde Fila Excel: Número de fila desde el que comienzan los datos.
            int Excel_Fila_Desde = 2; // la 1 es de la cabecera y la data comienza de las 2 en adelante.

            // Considera ultima columna (1=NO!, 0=Sí): La última columna es de observaciones, que no aplica para el Freemind.
            int Considera_Ultima_Columna = 1;

            // Nombre de la hoja que contiene la información en el Excel.
            String strNombreHoja = "GEC";
            #endregion

            // Genera archivo temporal 
            Procesa_Excel(Archivo_Excel, Archivo_Temp, Excel_Columna_Desde, Excel_Fila_Desde, Considera_Ultima_Columna, strNombreHoja);

            // Data desde: La primera columna del archivo temporal contiene datos. Este valor hace referencia a esa columna.
            int Data_Desde = 1;

            // Data Hasta: La última columna que pudiera contener datos. Este valor es variable en el archivo temporal
            // debido a que las celdas vacías no son recuperadas. Pero el valor máximo de columnas sí se sabe y este valor
            // se debe informar acá.
            int Data_Hasta = 8;

            // Procesa archivo y resultado lo deja en el portapapeles del sistema operativo.
            Procesa_Archivo_Temporal(Data_Desde, Data_Hasta, Archivo_Temp);

            Console.WriteLine("\nProceso finalizado.");
        }

        public static void Procesa_Excel(String Archivo_Excel, String Archivo_Temp, int Excel_Columna_Desde, int Excel_Fila_Desde, int Considera_Ultima_Columna , String strNombreHoja)
        {

            Excel.Application excelApp = new Excel.Application(); // Variable que corresponde con la app Excel.
            Excel.Workbook Workbooks = excelApp.Workbooks.Open(Archivo_Excel); // Abre el archivo Excel.
            Excel.Worksheet Hoja; // Variable que contendrá la hoja de Excel a procesar.
            Excel.Range data; // Variable que contendrá la data de la hoja de Excela que se procesará.

            Hoja = (Excel.Worksheet)Workbooks.Worksheets[strNombreHoja]; // Accede a la hoja que contiene los datos.
            data = Hoja.UsedRange; // Variable data queda con los datos que existen en la hoja.


            string strdato; // Variable que tomara dato de la celda que se este procesando.
            int rCnt = 0; // Variable asociada a fila (Row) que contiene información en la hoja de Excel.
            int cCnt = 0; // Variable asociada a la columna (Column) que contiene información en la hoja de Excel.

            String strLinea = ""; // String donde se irá almacenando la información de cada fila separa por un tabulador.
            String strTodaLaInfo = ""; //String que contendrá toda la información recuperada del excel.

            for (rCnt = Excel_Fila_Desde; rCnt <= data.Rows.Count; rCnt++) // Ciclo for para recorrer por filas.
            {
                strLinea = string.Empty; // Cada vez que comience una nueva columna, se limpia este String, para que almacene solo lo de la linea a procesar.
                for (cCnt = Excel_Columna_Desde; cCnt <= data.Columns.Count - Considera_Ultima_Columna; cCnt++) // ciclo for para recorrer cada columna de la fila que se está procesando.
                {
                    // Chiche para dar la impresión de que hace algo impresionante.
                    Console.Write(".");

                    strdato = (string)(data.Cells[rCnt, cCnt] as Excel.Range).Text; // obtiene dato de la celda.

                    if (String.IsNullOrEmpty(strLinea) ) // si está vacío, parte con el dato.
                    {
                        strLinea += strdato;
                    }
                    else
                    {
                        strLinea += "\t" + strdato; // Si ya tiene data, parte con un tabulador y le agrega el dato.
                    }
                }
                strTodaLaInfo += strLinea + "\n"; // Cuando termina de procesar todas las columnas, le agrega lo recolectado y le da un "Enter".
            }

            File.WriteAllText(Archivo_Temp, strTodaLaInfo); // una vez finalizado de recorrer la hoja con datos del excel, graba el resultado en un archivo temporal.

            // Vamoh a cerrarloh
            excelApp.DisplayAlerts = false;
            Workbooks.Close();
            excelApp.Quit();

            if (data != null)
            {
                Marshal.ReleaseComObject(data);
            }

            if (Hoja != null)
            {
                Marshal.ReleaseComObject(Hoja);
            }

            if (Workbooks != null)
            {
                Marshal.ReleaseComObject(Workbooks);
            }

            if (excelApp != null)
            {
                Marshal.ReleaseComObject(excelApp);
            }

        }

        public static void Procesa_Archivo_Temporal(int Data_Desde, int Data_Hasta, String Archivo_Temp)
        {
            String strTotalTexto = ""; // Variable que contendrá toda la información recolectada.

            string strlinea; // Variable que se asignará con la línea leida del archivo de texto temporal.
            List<String> CortesVariables = new List<string>(); // Lista que permitirá flexibilizar la cantidad de columnas variables que pudiera tener futuros archivos Excel.

            for (int intContador = 0; intContador <= (Data_Hasta - Data_Desde); intContador++) // Agrega data vacía según la cantidad máxima de columnas que pueda contener el archivo excel a procesar.
            {
                CortesVariables.Add(string.Empty);
            }

            String[] Cortes = CortesVariables.ToArray(); // Transforma la lista a un arreglo para poder realizar los corte de control.

            using (StreamReader file = new StreamReader(Archivo_Temp)) // ciclo de lectura del archivo temporal.
            {
                while ((strlinea = file.ReadLine()) != null) // Ciclo de lectura del archivo, línea por línea.
                {

                    char[] Delimitador_TAB = new char[] { '\t' }; // delimitador TAB
                    string[] strPartes = strlinea.Split(Delimitador_TAB, StringSplitOptions.RemoveEmptyEntries);  // almacena linea del archivo en un array por partes y quita las celdas vacías que venían en el Excel.

                    for (int i = 0; i < strPartes.Length; i++) // ciclo que recorre toda la fila, que puede variar en distintas filas.
                    {
                        string strTabs = new string('\t', i); // genera tantos TAB según la cantidad que indica la columna que está procesando.

                        if (Cortes[i] != strPartes[i]) // Evalúa corte de control. Si los datos son distintos, es que se cambió de rama.
                        {
                            // distinto
                            if (String.IsNullOrEmpty(strTotalTexto)) // Si es el primer dato, lo asigna directamente
                            {
                                strTotalTexto = strTabs + strPartes[i];
                            }
                            else
                            {
                                strTotalTexto += "\n" + strTabs + strPartes[i]; // Si ya tiene datos, le asigna una nueva línea y el dato de la nueva rama.
                            }
                            
                            Cortes[i] = strPartes[i]; // Asigna valor del dato, para que al volver a evaluar, vea si hay cambio de control o no.
                        }
                        else
                        {
                            //igual
                            strTotalTexto += strTabs; // Si está dentro la misma rama, solo asigna los TAB
                        }
                    }
                }
                file.Close(); // cierra el arrchivo abierto.
            }

            Clipboard.SetText(strTotalTexto); // Copia datos recopilados y formateados con la estructura al portapapeles.
            File.Delete(Archivo_Temp); // Borra archivo temporal.
        }
    }
}
