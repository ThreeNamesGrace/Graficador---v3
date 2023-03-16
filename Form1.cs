using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;

using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using System.Globalization;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Runtime.Versioning;
using System.Security;
using Microsoft.Win32;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////////////////////////////////////////////////////////

//////////////////////////////////////////////////////////////////////////////////////////////////////////////
namespace Graficador___v3
{       

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            //Form1.ShowWindow();
        }

        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        // Find window by Caption only. Note you must pass IntPtr.Zero as the first parameter.

        [DllImport("user32.dll", EntryPoint = "FindWindow", SetLastError = true)]
        static extern IntPtr FindWindowByCaption(IntPtr ZeroOnly, string lpWindowName);
       
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////////////////////////////////

        public Excel._Application GetOpenedExcelApplication()
        {
            Excel._Application objeto = null; //Declaramos un objeto para guardar la interfaz COM; lo declaramos como nulo para no tener problemas con el retorno.

            try //Intentamos...
            {
                objeto = (Excel._Application)ExMarshal.GetActiveObject("Excel.Application.16"); //Obtener el objeto activo y guardarlo.
                //Obtenemos un objeto "Excel.Application" que lo casteamos a un "Excel._Application".
            }
            catch (Exception ex) //Si ocurre un error atrapamos la excepción.
            {
                MessageBox.Show(ex.ToString()); //Mostramos la excepción que se presenta.
            }
            
            return objeto; //Regresamos el objeto COM.
        }

        ///----------------------------------------------------------------------------------------------------------------------
        private void GraficadorFrecuenciaGanancia_Click(object sender, EventArgs e)
        {
            //Código para saber el ProgID de Excel

            /*
            var regClis = Registry.ClassesRoot.OpenSubKey("CLSID");
            var progs = new List<string>();

            foreach (var clsid in regClis.GetSubKeyNames())
            {
                var regClsidKey = regClis.OpenSubKey(clsid);
                var ProgID = regClsidKey.OpenSubKey("ProgID");
                var regPath = regClsidKey.OpenSubKey("InprocServer32");

                if (regPath == null)
                    regPath = regClsidKey.OpenSubKey("LocalServer32");

                if (regPath != null && ProgID != null)
                {
                    var pid = ProgID.GetValue("");
                    var filePath = regPath.GetValue("");
                    progs.Add(pid + " -> " + filePath);
                    regPath.Close();
                }

                regClsidKey.Close();
            }
            foreach(var element in progs)
            {
                Debug.WriteLine(element);
            }
            */

            //Declaramos objetos que van a guardar los datos actuales del archivo de Excel.
            Excel._Application currentApplication = null;
            Excel.Workbooks currentWorkbooks = null;
            Excel._Workbook currentWorkbook = null;
            Excel.Sheets currentWorksheets = null;
            Excel.Worksheet currentWorksheet = null;
            Excel.Range allTheCells = null;

            //Declaramos las variables para guardar los rangos de las ganancias y las frecuencias.
            Excel.Range[] rangesOfGains = new Excel.Range[5];
            Excel.Range rangeOfFrequencies = null;

            //Celdas dummy.
            Excel.Range dummyCell1 = null;
            Excel.Range dummyCell2 = null;
            Excel.Range dummyCell3 = null;
            Excel.Range dummyCell4 = null;
                        
            //Variable para indicar que el sistema falló en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Checamos si Excel está abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicación de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que sí esté abierto.
            }
            catch (Exception ex) //Excel no está abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no esté abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos éxito.
                }
                
            }

            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colección de los libros de trabajo abiertos de la aplicación Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la única hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            while(currentWorksheets.Count > 1)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.
                
                currentWorksheets.Item[1].Delete(); //Borramos las demás hojas de trabajo con datos.
                                                    //Sólo debe de haber una hoja de trabajo.
            }

            ////Iteramos por cada hoja de gráfica y las borramos.
            //while(currentApplication.Charts.Count > 0)
            //{
            //    currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
            //    SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.
            //    currentApplication.DisplayAlerts = false; //Para que no tengamos que presionar Enter cada vez que se borra un ChartSheet.
            //    currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gráfica.
            //    currentApplication.DisplayAlerts = true; //Reactivamos las DialogBoxes.
                               
            //}

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Excel.Range titleCellDummy = null; //Celda dummy para verificar que existe el título.            
            System.Int32 titleCellRowNumber = 1;

            titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //La inicializamos en la mera primera celda.
            Boolean stillCount = true; //Ponemos una variable de control para seguir buscando tablas y aumentar la cuenta de ellas.

            int numberOfTables = 0;

            do //Usamos un "do" para que se ejecute el código al menos una vez.
            {
                if ((string)titleCellDummy.Value2 != null) //Vemos si hay texto en la celda (si hay texto es porque se hizo una tabla).
                {
                    numberOfTables++; //Aumentamos en uno el número de las tablas.
                    titleCellRowNumber += 54;
                    //titleCellDummy = null;
                    titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //Actualizamos las coordenadas de la celda de título.
                    stillCount = true; //Mantenemos la variable de control en verdadero.
                }
                else //Indicamos que hacer en caso de que no encuentre una tabla.
                {
                    stillCount = false; //Ponemos la variable de control en falso para detener la búsqueda (y por tanto la cuenta).
                }
            }
            while (stillCount == true); //Indicamos que siga buscando y contando mientras la variable de control sea verdadera.

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Para hacer la gráfica de ganancias.
            Excel.Chart[] GráficasFrecuenciaGanancia = new Excel.Chart[numberOfTables];

            //Para trabajar con cada gráfica.
            for (int i = 1; i <= numberOfTables; i++)
            {
                GráficasFrecuenciaGanancia[i - 1] = currentWorkbook.Charts.Add(); //Colocamos una hoja de trabajo para gráficas y la asignamos al objeto.
                               
                if (GráficasFrecuenciaGanancia[i-1] == null)
                {
                    MessageBox.Show("Error", "Null found");                    
                }                          
                

                GráficasFrecuenciaGanancia[i - 1].ChartWizard(
                    Gallery: XlChartType.xlLineMarkers,
                    PlotBy: XlRowCol.xlColumns,
                    CategoryLabels: 45,
                    SeriesLabels: 5,
                    HasLegend: true,
                    Title: allTheCells.Item[1 + ((i - 1) * 54), 1].Value2,
                    CategoryTitle: "Frecuencias (Hz)",
                    ValueTitle: "Ganancias (dB)",
                    ExtraTitle: "Extra"
                    );

                //foreach(Process proceso in ExcelProcesses)
                //{
                //    proceso.Kill();
                //}
                
                Excel.Characters TítuloGráfica = null; //Creamos un objecto Characters para ponerle título a la gráfica.
                string dummyString = null; //Creamos un string para guardar el título que le vamos a poner a la gráfica.
                dummyString = Convert.ToString(allTheCells.Item[1 + ((i - 1) * 54), 1].Value2); //Obtenenmos el título de la gráfica y lo guardamos en la variable temporal.
                dummyString = dummyString.Trim('\t', '\r', '\n'); //Quitamos los caracteres de escape que dan error.
                dummyString = dummyString.Substring(108);

                //Hacemos lo anterior antes de llenar la gráfica con datos.
                //Lo siguiente se explica por sí mismo.
                GráficasFrecuenciaGanancia[i - 1].HasLegend = true; //Activamos que tenga leyenda.
                GráficasFrecuenciaGanancia[i - 1].ChartWizard( //Utilizamos el ChartWizard para ayudarnos a hacer las gráficas.
                    Gallery: XlChartType.xlXYScatterLines,
                    PlotBy: XlRowCol.xlColumns,                    
                    SeriesLabels: 5,
                    HasLegend: true,
                    Title: dummyString,
                    CategoryTitle: "Frecuencias (Hz)",
                    ValueTitle: "Ganancias (dB)"
                    //ExtraTitle: "Extra"
                    );

                GráficasFrecuenciaGanancia[i - 1].ChartType = XlChartType.xlXYScatterLines;
                GráficasFrecuenciaGanancia[i - 1].HasTitle = true; //Activamos que tenga título.
                TítuloGráfica = GráficasFrecuenciaGanancia[i - 1].ChartTitle.Characters; //Pasamos el título de la gráfica al objeto indicado.
                currentWorkbook.Activate(); //Activamos el libro de trabajo.
                TítuloGráfica.Text = dummyString;

                //Asignamos el string a la propiedad Text; le ponemos el texto al...
                //título de la gráfica.

                //if (dummyString == null) //Checamos si el string que vamos a asignar no está en blanco.
                //{
                //    MessageBox.Show("Object is Null", "Null object"); //Mandamos un aviso de que la celda de la que se obtiene el string está vacía.
                //}

                //try
                //{


                //}
                //catch (COMException ex) //Atrapamos la excepción que sale.
                //{
                //    Debug.Write(ex.Message + "\n"); //Mostramos la excepción en la consola.
                //    Debug.WriteLine(ex.HResult.ToString("X")); //Mostramos su identificador HResult.
                //    Debug.WriteLine(dummyString); //Mostramos si el string está vacío o si estuvo vacío.
                //    Debug.WriteLine("");
                //}

                //Asignamos las celdas a las variables especificadas.
                //Para las ganancias:
                dummyCell1 = allTheCells.Item[4 + 54 * (i - 1), 2];
                dummyCell2 = allTheCells.Item[4 + 48 + 54 * (i - 1), 2];
                //Para las frecuencias:
                dummyCell3 = allTheCells.Item[4 + 54 * (i - 1), 1];
                dummyCell4 = allTheCells.Item[4 + 48 + 54 * (i - 1), 1];

                //Ciclo For para guardar los rangos de las celdas para las ganancias.
                for (int j = 0; j < 5; j++)
                {
                    rangesOfGains[j] = allTheCells.Range[dummyCell1, dummyCell2];
                    dummyCell1 = allTheCells.Item[dummyCell1.Row, dummyCell1.Column + 4];
                    dummyCell2 = allTheCells.Item[dummyCell2.Row, dummyCell2.Column + 4];
                }

                //Creamos el rango de las celdas que contienen las frecuencias.
                rangeOfFrequencies = allTheCells.Range[dummyCell3, dummyCell4];

                //Hacemos las series de datos
                Excel.SeriesCollection[] ColecciónDeSeries = new Excel.SeriesCollection[numberOfTables]; //Hacemos un array de colecciones de series, en el que cada elemento...
                //es la colección de series de una gráfica en particular.
                Excel.Series[] SeriesDeDatos = new Excel.Series[5 * numberOfTables];
                ColecciónDeSeries[i - 1] = GráficasFrecuenciaGanancia[i - 1].SeriesCollection(); //Obtenenmos la colección de series de la gráfica.

                while(ColecciónDeSeries[i-1].Count > 0) //Mientras haya series en la colección de series de una gráfica.
                {
                    ColecciónDeSeries[i - 1].Item(1).Delete();  //Borramos cada series.
                    //Cada vez que se borra una series, la siguiete serie se vuelve la primera.
                    //Borramos todas las series porque la gráfica debe estar vacía desde su creación.                                                                
                }
                List<String> frecuenciasEnStrings = new List<String>(); //Creamos una lista para guardar las frecuencias para hacer la gráfica logaritmica.
                                                                          
                foreach(Excel.Range celda in rangeOfFrequencies) //Para obtener los valores de las celdas y convertirlas en strings.
                {
                    frecuenciasEnStrings.Add(Convert.ToString(celda.Value2)); //Agregamos cada valor de frecuencia en la lista de strings.
                    //Debug.WriteLine(frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("k"));
                    if(frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] == null) //Vemos si el elemento no es nulo
                    {
                        break;
                    }
                    else
                    {
                        if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("k")) //Buscamos si el último string contiene kilo.
                        {
                            frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] = frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Replace("k", "000"); //Reemplazamos "k" por los ceros que representa.
                        }
                        if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("M")) //Buscamos si el último string contiene mega.
                        {
                            frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] = frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Replace("M", "000000"); //Reemplazamos "M" por los ceros que representa.
                        }
                    }                    
                }
                if(frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] == null)
                {
                    frecuenciasEnStrings.RemoveAt(frecuenciasEnStrings.Count() - 1);
                }

                List<int> frecuenciasEnEnteros = new List<int>(); //Creamos una lista de números enteros que son las frecuencias.
                foreach(string elemento in frecuenciasEnStrings) //Para convertir los strings anteriores en números enteros.
                {
                    frecuenciasEnEnteros.Add(Convert.ToInt32(elemento)); //Convertimos cada string en un entero.
                }

                int[] frecuencias = frecuenciasEnEnteros.ToArray(); 
                for (int j = 5 * (i - 1); j < 5 * (i); j++) //Ciclamos en las 5 mediciones.
                {                                 
                    SeriesDeDatos[j] = ColecciónDeSeries[i - 1].NewSeries(); //Creamos una nueva serie en la gráfica.
                    //SeriesDeDatos[j].ClearFormats();
                    SeriesDeDatos[j].XValues = frecuencias; //Introducimos el rango de celdas donde se encuentran las frecuencias.
                    SeriesDeDatos[j].Values = rangesOfGains[j % 5]; //Introducimos el rango de celdas donde se encuentran las ganancias.
                    //Usamos modulo para sacar valores desde cero a 5.
                    SeriesDeDatos[j].Name = "Medición #" + ((j % 5) + 1).ToString(); //Le ponemos nombre a cada serie.
                }

                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                ////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                
                
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
                
                //Modificamos los ejes.
                Excel.Axes AxesOfTheChart = null; //Creamos un objeto de colección de ejes.
                AxesOfTheChart = GráficasFrecuenciaGanancia[i - 1].Axes(); //Enlazamos el objeto con los ejes de la gráfica.
                Excel.Axis XAxis = AxesOfTheChart.Item(XlAxisType.xlCategory, XlAxisGroup.xlPrimary); //Obtenenmos el eje "X" primario.
                XAxis.HasMajorGridlines = true; //Activamos las líneas de cuadrícula mayores.
                XAxis.HasMinorGridlines = true; //Activamos las líneas de cuadrícula menores.
                //xis.AxisBetweenCategories = false;
                
                //Hacemos los mismo para el eje "Y".
                Excel.Axis YAxis = AxesOfTheChart.Item(XlAxisType.xlValue, XlAxisGroup.xlPrimary); //Debemos sacar los primarios, porque son los que se pueden modificar.
                YAxis.HasMajorGridlines = true; //Activamos líneas de cuadrícula mayores.
                YAxis.HasMinorGridlines = true; //Activamos líneas de cuadrícula menores.

                //Modificamos las líneas de cuadrícula del eje "X".
                Excel.Gridlines LíneasMenoresEjeX = null; //Creamos el objeto interfaz.
                LíneasMenoresEjeX = XAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMenoresEjeX = LíneasMenoresEjeX.Border; //Creamos un objeto borde y lo enlazamos.
                BordeLíneasMenoresEjeX.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un número hexadecimal.

                Excel.Gridlines LíneasMayoresEjeX = null; //Creamos el objeto interfaz.
                LíneasMayoresEjeX = XAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMayoresEjeX = LíneasMayoresEjeX.Border; //Creamos el objeto borde y lo enlazamos.
                BordeLíneasMayoresEjeX.Color = 0xD9D9D9; //Modificamos su color.

                //Modificamos las líneas de cuadrícula del eje "Y".
                Excel.Gridlines LíneasMenoresEjeY = null; //Creamos el objeto interfaz.
                LíneasMenoresEjeY = YAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMenoresEjeY = LíneasMenoresEjeY.Border; //Creamos un objeto borde y lo enlazamos.
                BordeLíneasMenoresEjeY.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un número hexadecimal.

                Excel.Gridlines LíneasMayoresEjeY = null; //Creamos el objeto interfaz.
                LíneasMayoresEjeY = YAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMayoresEjeY = LíneasMayoresEjeY.Border; //Creamos el objeto borde y lo enlazamos.
                BordeLíneasMayoresEjeY.Color = 0xD9D9D9; //Modificamos su color.

                //para hacer una gráfica semilogarítmica.
                XAxis.ScaleType = XlScaleType.xlScaleLogarithmic; //Hacemos que el eje "X" sea logaritmico
                XAxis.LogBase = (double)10; //Ponemos la base del logaritmo.
                XAxis.MinimumScale = 200; //Ponemos el valor mínimo del eje "X".
                XAxis.MaximumScale = 20000000; //Ponemos el valor mínimo del eje "X".               
                YAxis.MinimumScale = -40; //Ponemos el valor mínimo del eje "Y".
                YAxis.MaximumScale = 0; //Ponemos el valor máximo del eje "Y".
                XAxis.DisplayUnit = XlDisplayUnit.xlThousands; //Que las etiquetas del eje "X" sean múltiplos de 1000.

                GráficasFrecuenciaGanancia[i - 1].PlotArea.Width = 500; //Que el ancho del gráfico sea de 500 puntos.
                GráficasFrecuenciaGanancia[i - 1].PlotArea.Height = 500; //Que el alto del gráfico sea de 500 puntos.
            }

            Form Form1 = this; //Creamos un objeto Form y le asignamos la ventana de la form con los botones.
            Form1.Show(); //La mostramos después de que ejecuta todo el código de arriba.
        }


        private void GráficaFrecuenciaFase_Click(object sender, EventArgs e)
        {
            //Declaramos objetos que van a guardar los datos actuales del archivo de Excel.
            Excel._Application currentApplication = null;
            Excel.Workbooks currentWorkbooks = null;
            Excel._Workbook currentWorkbook = null;
            Excel.Sheets currentWorksheets = null;
            Excel.Worksheet currentWorksheet = null;
            Excel.Range allTheCells = null;

            //Declaramos las variables para guardar los rangos de las ganancias y las frecuencias.
            Excel.Range[] rangesOfPhase = new Excel.Range[5];
            Excel.Range rangeOfFrequencies = null;

            //Celdas dummy.
            Excel.Range dummyCell1 = null;
            Excel.Range dummyCell2 = null;
            Excel.Range dummyCell3 = null;
            Excel.Range dummyCell4 = null;
                        
            //Variable para indicar que el sistema falló en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            //Checamos si Excel está abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicación de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que sí esté abierto.
            }
            catch (Exception ex) //Excel no está abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no esté abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos éxito.
                }

            }
            
            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colección de los libros de trabajo abiertos de la aplicación Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la única hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            //while (currentWorksheets.Count > 1)
            //{
            //    currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
            //    SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.

            //    currentWorksheets.Item[1].Delete(); //Borramos las demás hojas de trabajo con datos.
            //                                        //Sólo debe de haber una hoja de trabajo.
            //}

            ////Iteramos por cada hoja de gráfica y las borramos.
            //while (currentApplication.Charts.Count > 0)
            //{
            //    currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
            //    SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.
            //    currentApplication.DisplayAlerts = false;
            //    currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gráfica.
            //    currentApplication.DisplayAlerts = true;
            //}

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Excel.Range titleCellDummy = null; //Celda dummy para verificar que existe el título.
            System.Int32 titleCellRowNumber = 1;

            titleCellDummy = allTheCells.Item[1, 1]; //La inicializamos en la mera primera celda.
            Boolean stillCount = true; //Ponemos una variable de control para seguir buscando tablas y aumentar la cuenta de ellas.

            int numberOfTables = 0;

            do //Usamos un "do" para que se ejecute el código al menos una vez.
            {
                if (titleCellDummy.Value2 != null) //Vemos si hay texto en la celda (si hay texto es porque se hizo una tabla).
                {
                    numberOfTables++; //Aumentamos en uno el número de las tablas.
                    titleCellRowNumber += 54;
                    //titleCellDummy = null;
                    titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //Actualizamos las coordenadas de la celda de título.
                    stillCount = true; //Mantenemos la variable de control en verdadero.
                }
                else //Indicamos que hacer en caso de que no encuentre una tabla.
                {
                    stillCount = false; //Ponemos la variable de control en falso para detener la búsqueda (y por tanto la cuenta).
                }
            }
            while (stillCount == true); //Indicamos que siga buscando y contando mientras la variable de control sea verdadera.

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Para hacer la gráfica de ganancias.
            Excel.Chart[] GráficasFrecuenciaFase = new Excel.Chart[numberOfTables];

            //Para trabajar con cada gráfica.
            for (int i = 1; i <= numberOfTables; i++)
            {
                GráficasFrecuenciaFase[i - 1] = currentWorkbook.Charts.Add(); //Colocamos una hoja de trabajo para gráficas y la asignamos al objeto.

                Excel.Characters TítuloGráfica = null; //Creamos un objecto Characters para ponerle título a la gráfica.
                string dummyString = null; //Creamos un string para guardar el título que le vamos a poner a la gráfica.
                dummyString = Convert.ToString(allTheCells.Item[1 + ((i - 1) * 54), 1].Value2); //Obtenenmos el título de la gráfica y lo guardamos en la variable temporal.
                dummyString = dummyString.Trim('\t', '\r', '\n'); //Quitamos los caracteres de escape que dan error.
                dummyString = dummyString.Substring(108);

                //Hacemos lo anterior antes de llenar la gráfica con datos.
                //Lo siguiente se explica por sí mismo.
                GráficasFrecuenciaFase[i - 1].HasLegend = true; //Activamos que tenga leyenda.
                GráficasFrecuenciaFase[i - 1].ChartWizard( //Utilizamos el ChartWizard para ayudarnos a hacer las gráficas.
                    Gallery: XlChartType.xlXYScatterLines,
                    PlotBy: XlRowCol.xlColumns,
                    SeriesLabels: 5,
                    HasLegend: true,
                    Title: dummyString,
                    CategoryTitle: "Frecuencias (Hz)",
                    ValueTitle: "Ganancias (dB)"
                    //ExtraTitle: "Extra"
                    );

                GráficasFrecuenciaFase[i - 1].ChartType = XlChartType.xlXYScatterLines;
                GráficasFrecuenciaFase[i - 1].HasTitle = true; //Activamos que tenga título.
                TítuloGráfica = GráficasFrecuenciaFase[i - 1].ChartTitle.Characters; //Pasamos el título de la gráfica al objeto indicado.
                currentWorkbook.Activate(); //Activamos el libro de trabajo.
                TítuloGráfica.Text = dummyString;

                //Asignamos las celdas a las variables especificadas.
                //Para las ganancias:
                dummyCell1 = allTheCells.Item[4 + 54 * (i - 1), 4];
                dummyCell2 = allTheCells.Item[4 + 28 + 54 * (i - 1), 4];
                //Para las frecuencias:
                dummyCell3 = allTheCells.Item[4 + 54 * (i - 1), 1];
                dummyCell4 = allTheCells.Item[4 + 28 + 54 * (i - 1), 1];

                //Ciclo For para guardar los rangos de las celdas para las ganancias.
                for (int j = 0; j < 5; j++)
                {
                    rangesOfPhase[j] = allTheCells.Range[dummyCell1, dummyCell2];
                    dummyCell1 = allTheCells.Item[dummyCell1.Row, dummyCell1.Column + 4];
                    dummyCell2 = allTheCells.Item[dummyCell2.Row, dummyCell2.Column + 4];
                }

                //Creamos el rango de las celdas que contienen las frecuencias.
                rangeOfFrequencies = allTheCells.Range[dummyCell3, dummyCell4];

                //Hacemos las series de datos
                Excel.SeriesCollection[] ColecciónDeSeries = new Excel.SeriesCollection[numberOfTables]; //Hacemos un array de colecciones de series, en el que cada elemento...
                //es la colección de series de una gráfica en particular.
                Excel.Series[] SeriesDeDatos = new Excel.Series[5 * numberOfTables];
                ColecciónDeSeries[i - 1] = GráficasFrecuenciaFase[i - 1].SeriesCollection(); //Obtenenmos la colección de series de la gráfica.

                while (ColecciónDeSeries[i - 1].Count > 0) //Mientras haya series en la colección de series de una gráfica.
                {
                    ColecciónDeSeries[i - 1].Item(1).Delete();  //Borramos cada series.
                                                                //Cada vez que se borra una series, la siguiete serie se vuelve la primera.
                                                                //Borramos todas las series porque la gráfica debe estar vacía desde su creación.

                }

                List<String> frecuenciasEnStrings = new List<String>(); //Creamos una lista para guardar las frecuencias para hacer la gráfica logaritmica.

                foreach (Excel.Range celda in rangeOfFrequencies) //Para obtener los valores de las celdas y convertirlas en strings.
                {
                    frecuenciasEnStrings.Add(Convert.ToString(celda.Value2)); //Agregamos cada valor de frecuencia en la lista de strings.
                    //Debug.WriteLine(frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("k"));
                    if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] == null) //Vemos si el elemento no es nulo
                    {
                        break;
                    }
                    else
                    {
                        if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("k")) //Buscamos si el último string contiene kilo.
                        {
                            frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] = frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Replace("k", "000"); //Reemplazamos "k" por los ceros que representa.
                        }
                        if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Contains("M")) //Buscamos si el último string contiene mega.
                        {
                            frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] = frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1].Replace("M", "000000"); //Reemplazamos "M" por los ceros que representa.
                        }
                    }
                }
                if (frecuenciasEnStrings[frecuenciasEnStrings.Count() - 1] == null)
                {
                    frecuenciasEnStrings.RemoveAt(frecuenciasEnStrings.Count() - 1);
                }

                List<int> frecuenciasEnEnteros = new List<int>(); //Creamos una lista de números enteros que son las frecuencias.
                foreach (string elemento in frecuenciasEnStrings) //Para convertir los strings anteriores en números enteros.
                {
                    frecuenciasEnEnteros.Add(Convert.ToInt32(elemento)); //Convertimos cada string en un entero.
                }

                int[] frecuencias = frecuenciasEnEnteros.ToArray();

                for (int j = 5 * (i - 1); j < 5 * (i); j++) //Ciclamos en las 5 mediciones.
                {
                    SeriesDeDatos[j] = ColecciónDeSeries[i - 1].NewSeries(); //Creamos una nueva serie en la gráfica.
                    //SeriesDeDatos[j].ClearFormats();
                    SeriesDeDatos[j].XValues = frecuencias; //Introducimos el rango de celdas donde se encuentran las frecuencias.
                    SeriesDeDatos[j].Values = rangesOfPhase[j % 5]; //Introducimos el rango de celdas donde se encuentran las ganancias.
                    //Usamos modulo para sacar valores desde cero a 5.
                    SeriesDeDatos[j].Name = "Medición #" + ((j % 5) + 1).ToString(); //Le ponemos nombre a cada serie.
                }

                //Modificamos los ejes.
                Excel.Axes AxesOfTheChart = null; //Creamos un objeto de colección de ejes.
                AxesOfTheChart = GráficasFrecuenciaFase[i - 1].Axes(); //Enlazamos el objeto con los ejes de la gráfica.
                Excel.Axis XAxis = AxesOfTheChart.Item(XlAxisType.xlCategory, XlAxisGroup.xlPrimary); //Obtenenmos el eje "X" primario.
                XAxis.HasMajorGridlines = true; //Activamos las líneas de cuadrícula mayores.
                XAxis.HasMinorGridlines = true; //Activamos las líneas de cuadrícula menores.
                
                //Hacemos los mismo para el eje "Y".
                Excel.Axis YAxis = AxesOfTheChart.Item(XlAxisType.xlValue, XlAxisGroup.xlPrimary); //Debemos sacar los primarios, porque son los que se pueden modificar.
                YAxis.HasMajorGridlines = true; //Activamos líneas de cuadrícula mayores.
                YAxis.HasMinorGridlines = true; //Activamos líneas de cuadrícula menores.

                //Modificamos las líneas de cuadrícula del eje "X".
                Excel.Gridlines LíneasMenoresEjeX = null; //Creamos el objeto interfaz.
                LíneasMenoresEjeX = XAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMenoresEjeX = LíneasMenoresEjeX.Border; //Creamos un objeto borde y lo enlazamos.
                BordeLíneasMenoresEjeX.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un número hexadecimal.

                Excel.Gridlines LíneasMayoresEjeX = null; //Creamos el objeto interfaz.
                LíneasMayoresEjeX = XAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMayoresEjeX = LíneasMayoresEjeX.Border; //Creamos el objeto borde y lo enlazamos.
                BordeLíneasMayoresEjeX.Color = 0xD9D9D9; //Modificamos su color.

                //Modificamos las líneas de cuadrícula del eje "Y".
                Excel.Gridlines LíneasMenoresEjeY = null; //Creamos el objeto interfaz.
                LíneasMenoresEjeY = YAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMenoresEjeY = LíneasMenoresEjeY.Border; //Creamos un objeto borde y lo enlazamos.
                BordeLíneasMenoresEjeY.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un número hexadecimal.

                Excel.Gridlines LíneasMayoresEjeY = null; //Creamos el objeto interfaz.
                LíneasMayoresEjeY = YAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeLíneasMayoresEjeY = LíneasMayoresEjeY.Border; //Creamos el objeto borde y lo enlazamos.
                BordeLíneasMayoresEjeY.Color = 0xD9D9D9; //Modificamos su color.

                //para hacer una gráfica semilogarítmica.
                XAxis.ScaleType = XlScaleType.xlScaleLogarithmic; //Hacemos que el eje "X" sea logaritmico
                XAxis.LogBase = (double)10; //Ponemos la base del logaritmo.
                XAxis.MinimumScale = 200; //Ponemos el valor mínimo del eje "X".
                XAxis.MaximumScale = 2000000; //Ponemos el valor mínimo del eje "X".
                YAxis.MinimumScale = -60; //Ponemos el valor mínimo del eje "Y".
                YAxis.MaximumScale = 0; //Ponemos el valor máximo del eje "Y".
                XAxis.DisplayUnit = XlDisplayUnit.xlThousands; //Que las etiquetas del eje "X" sean múltiplos de 1000.

                GráficasFrecuenciaFase[i - 1].PlotArea.Width = 500; //Que el ancho del área de la gráfica sea de 500 puntos.
                GráficasFrecuenciaFase[i - 1].PlotArea.Height = 500; //Que el ancho del área de la gráfica sea de 500 puntos.                
            }

            Form Form1 = this; //Creamos un objeto Form y le asignamos la ventana de la form con los botones.
            Form1.Show(); //La mostramos después de que ejecuta todo el código de arriba.
        }

        private void BorraGráficas_Click(object sender, EventArgs e)
        {
            //Declaramos objetos que van a guardar los datos actuales del archivo de Excel.
            Excel._Application currentApplication = null;
            Excel.Workbooks currentWorkbooks = null;
            Excel._Workbook currentWorkbook = null;
            Excel.Sheets currentWorksheets = null;
            Excel.Worksheet currentWorksheet = null;
            Excel.Range allTheCells = null;
                       
            //Variable para indicar que el sistema falló en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            //Checamos si Excel está abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicación de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que sí esté abierto.
            }
            catch (Exception ex) //Excel no está abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no esté abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos éxito.
                }

            }

            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colección de los libros de trabajo abiertos de la aplicación Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la única hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            while (currentWorksheets.Count > 1)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.

                currentWorksheets.Item[1].Delete(); //Borramos las demás hojas de trabajo con datos.
                                                    //Sólo debe de haber una hoja de trabajo.
            }

            //Iteramos por cada hoja de gráfica y las borramos.
            while (currentApplication.Charts.Count > 0)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicación de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicación esté al frente.

                currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gráfica.

            }

            Form Form1 = this; //Creamos un objeto Form y le asignamos la ventana de la form con los botones.
            Form1.Show(); //La mostramos después de que ejecuta todo el código de arriba.
        }

        private void CerrarPrograma_Click(object sender, EventArgs e)
        {
            Form Form1 = this; //Creamos una instancia Form y le asignamos la ventana de la Form
            Form1.Close(); //Cerrarmos la aplicación Windows Form.
        }
    }

    //Creamos una clase con código antiguo de Marshal que es necesario para este programa.
    public static class ExMarshal
    {
        internal const String OLEAUT32 = "oleaut32.dll";
        internal const String OLE32 = "ole32.dll";

        [System.Security.SecurityCritical]  // auto-generated_required
        public static Object GetActiveObject(String progID) //Nos permite obtener un objeto COM.
        {
            Object obj = null;
            Guid clsid;

            // Call CLSIDFromProgIDEx first then fall back on CLSIDFromProgID if
            // CLSIDFromProgIDEx doesn't exist.
            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            //            catch
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLE32, PreserveSig = false)]
        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        //[DllImport(Microsoft.Win32.Win32Native.OLEAUT32, PreserveSig = false)]
        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]  // auto-generated
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);
    }
}




