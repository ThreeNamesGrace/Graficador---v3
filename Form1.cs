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
            catch (Exception ex) //Si ocurre un error atrapamos la excepci�n.
            {
                MessageBox.Show(ex.ToString()); //Mostramos la excepci�n que se presenta.
            }
            
            return objeto; //Regresamos el objeto COM.
        }

        ///----------------------------------------------------------------------------------------------------------------------
        private void GraficadorFrecuenciaGanancia_Click(object sender, EventArgs e)
        {
            //C�digo para saber el ProgID de Excel

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
                        
            //Variable para indicar que el sistema fall� en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Checamos si Excel est� abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicaci�n de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que s� est� abierto.
            }
            catch (Exception ex) //Excel no est� abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no est� abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos �xito.
                }
                
            }

            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colecci�n de los libros de trabajo abiertos de la aplicaci�n Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la �nica hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            while(currentWorksheets.Count > 1)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.
                
                currentWorksheets.Item[1].Delete(); //Borramos las dem�s hojas de trabajo con datos.
                                                    //S�lo debe de haber una hoja de trabajo.
            }

            //Iteramos por cada hoja de gr�fica y las borramos.
            while(currentApplication.Charts.Count > 0)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.
                
                currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gr�fica.
                
            }

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Excel.Range titleCellDummy = null; //Celda dummy para verificar que existe el t�tulo.            
            System.Int32 titleCellRowNumber = 1;

            titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //La inicializamos en la mera primera celda.
            Boolean stillCount = true; //Ponemos una variable de control para seguir buscando tablas y aumentar la cuenta de ellas.

            int numberOfTables = 0;

            do //Usamos un "do" para que se ejecute el c�digo al menos una vez.
            {
                if ((string)titleCellDummy.Value2 != null) //Vemos si hay texto en la celda (si hay texto es porque se hizo una tabla).
                {
                    numberOfTables++; //Aumentamos en uno el n�mero de las tablas.
                    titleCellRowNumber += 54;
                    //titleCellDummy = null;
                    titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //Actualizamos las coordenadas de la celda de t�tulo.
                    stillCount = true; //Mantenemos la variable de control en verdadero.
                }
                else //Indicamos que hacer en caso de que no encuentre una tabla.
                {
                    stillCount = false; //Ponemos la variable de control en falso para detener la b�squeda (y por tanto la cuenta).
                }
            }
            while (stillCount == true); //Indicamos que siga buscando y contando mientras la variable de control sea verdadera.

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Para hacer la gr�fica de ganancias.
            Excel.Chart[] Gr�ficasFrecuenciaGanancia = new Excel.Chart[numberOfTables];

            //Para trabajar con cada gr�fica.
            for (int i = 1; i <= numberOfTables; i++)
            {
                Gr�ficasFrecuenciaGanancia[i - 1] = currentWorkbook.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing); //Colocamos una hoja de trabajo para gr�ficas y la asignamos al objeto.
                               
                if (Gr�ficasFrecuenciaGanancia[i-1] == null)
                {
                    MessageBox.Show("Error", "Null found");                    
                }                          
                

                Gr�ficasFrecuenciaGanancia[i - 1].ChartWizard(
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

                Gr�ficasFrecuenciaGanancia[i - 1].ChartType = XlChartType.xlLineMarkers;
                Gr�ficasFrecuenciaGanancia[i-1].HasTitle = true;
                //Excel.ChartTitle T�tuloGr�fica = null;
                //T�tuloGr�fica = Gr�ficasFrecuenciaGanancia[i - 1].ChartTitle;
                //T�tuloGr�fica.Text = allTheCells.Item[1 + ((i - 1) * 54), 1].Value2;

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
                Excel.SeriesCollection[] Colecci�nDeSeries = new Excel.SeriesCollection[numberOfTables]; //Hacemos un array de colecciones de series, en el que cada elemento...
                //es la colecci�n de series de una gr�fica en particular.
                Excel.Series[] SeriesDeDatos = new Excel.Series[5 * numberOfTables];
                Colecci�nDeSeries[i - 1] = Gr�ficasFrecuenciaGanancia[i - 1].SeriesCollection(); //Obtenenmos la colecci�n de series de la gr�fica.

                while(Colecci�nDeSeries[i-1].Count > 0) //Mientras haya series en la colecci�n de series de una gr�fica.
                {
                    Colecci�nDeSeries[i - 1].Item(1).Delete();  //Borramos cada series.
                    //Cada vez que se borra una series, la siguiete serie se vuelve la primera.
                    //Borramos todas las series porque la gr�fica debe estar vac�a desde su creaci�n.
                                                                
                }
                                             

                for (int j = 5 * (i - 1); j < 5 * (i); j++) //Ciclamos en las 5 mediciones.
                {                                 
                    SeriesDeDatos[j] = Colecci�nDeSeries[i - 1].NewSeries(); //Creamos una nueva serie en la gr�fica.
                    //SeriesDeDatos[j].ClearFormats();
                    SeriesDeDatos[j].XValues = rangeOfFrequencies.Cells; //Introducimos el rango de celdas donde se encuentran las frecuencias.
                    SeriesDeDatos[j].Values = rangesOfGains[j % 5]; //Introducimos el rango de celdas donde se encuentran las ganancias.
                    //Usamos modulo para sacar valores desde cero a 5.
                    SeriesDeDatos[j].Name = "Medici�n #" + ((j % 5) + 1).ToString(); //Le ponemos nombre a cada serie.
                }

                //Modificamos los ejes.
                Excel.Axes AxesOfTheChart = null; //Creamos un objeto de colecci�n de ejes.
                AxesOfTheChart = Gr�ficasFrecuenciaGanancia[i - 1].Axes(); //Enlazamos el objeto con los ejes de la gr�fica.
                Excel.Axis XAxis = AxesOfTheChart.Item(XlAxisType.xlCategory, XlAxisGroup.xlPrimary); //Obtenenmos el eje "X" primario.
                XAxis.HasMajorGridlines = true; //Activamos las l�neas de cuadr�cula mayores.
                XAxis.HasMinorGridlines = true; //Activamos las l�neas de cuadr�cula menores.
                XAxis.AxisBetweenCategories = false;

                //Hacemos los mismo para el eje "Y".
                Excel.Axis YAxis = AxesOfTheChart.Item(XlAxisType.xlValue, XlAxisGroup.xlPrimary); //Debemos sacar los primarios, porque son los que se pueden modificar.
                YAxis.HasMajorGridlines = true; //Activamos l�neas de cuadr�cula mayores.
                YAxis.HasMinorGridlines = true; //Activamos l�neas de cuadr�cula menores.

                //Modificamos las l�neas de cuadr�cula del eje "X".
                Excel.Gridlines L�neasMenoresEjeX = null; //Creamos el objeto interfaz.
                L�neasMenoresEjeX = XAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMenoresEjeX = L�neasMenoresEjeX.Border; //Creamos un objeto borde y lo enlazamos.
                BordeL�neasMenoresEjeX.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un n�mero hexadecimal.

                Excel.Gridlines L�neasMayoresEjeX = null; //Creamos el objeto interfaz.
                L�neasMayoresEjeX = XAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMayoresEjeX = L�neasMayoresEjeX.Border; //Creamos el objeto borde y lo enlazamos.
                BordeL�neasMayoresEjeX.Color = 0xD9D9D9; //Modificamos su color.

                //Modificamos las l�neas de cuadr�cula del eje "Y".
                Excel.Gridlines L�neasMenoresEjeY = null; //Creamos el objeto interfaz.
                L�neasMenoresEjeY = YAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMenoresEjeY = L�neasMenoresEjeY.Border; //Creamos un objeto borde y lo enlazamos.
                BordeL�neasMenoresEjeY.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un n�mero hexadecimal.

                Excel.Gridlines L�neasMayoresEjeY = null; //Creamos el objeto interfaz.
                L�neasMayoresEjeY = YAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMayoresEjeY = L�neasMayoresEjeY.Border; //Creamos el objeto borde y lo enlazamos.
                BordeL�neasMayoresEjeY.Color = 0xD9D9D9; //Modificamos su color.

                //MessageBox.Show(String.Format("Chart {0}", currentWorkbook.Charts.Item[currentWorkbook.Charts.Count].Name));
            }


        }


        private void Gr�ficaFrecuenciaFase_Click(object sender, EventArgs e)
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
                        
            //Variable para indicar que el sistema fall� en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            //Checamos si Excel est� abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicaci�n de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que s� est� abierto.
            }
            catch (Exception ex) //Excel no est� abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no est� abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos �xito.
                }

            }
            
            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colecci�n de los libros de trabajo abiertos de la aplicaci�n Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la �nica hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            while (currentWorksheets.Count > 1)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.

                currentWorksheets.Item[1].Delete(); //Borramos las dem�s hojas de trabajo con datos.
                                                    //S�lo debe de haber una hoja de trabajo.
            }

            //Iteramos por cada hoja de gr�fica y las borramos.
            while (currentApplication.Charts.Count > 0)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.

                currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gr�fica.

            }

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            Excel.Range titleCellDummy = null; //Celda dummy para verificar que existe el t�tulo.
            System.Int32 titleCellRowNumber = 1;

            titleCellDummy = allTheCells.Item[1, 1]; //La inicializamos en la mera primera celda.
            Boolean stillCount = true; //Ponemos una variable de control para seguir buscando tablas y aumentar la cuenta de ellas.

            int numberOfTables = 0;

            do //Usamos un "do" para que se ejecute el c�digo al menos una vez.
            {
                if (titleCellDummy.Value2 != null) //Vemos si hay texto en la celda (si hay texto es porque se hizo una tabla).
                {
                    numberOfTables++; //Aumentamos en uno el n�mero de las tablas.
                    titleCellRowNumber += 54;
                    //titleCellDummy = null;
                    titleCellDummy = allTheCells.Item[titleCellRowNumber, 1]; //Actualizamos las coordenadas de la celda de t�tulo.
                    stillCount = true; //Mantenemos la variable de control en verdadero.
                }
                else //Indicamos que hacer en caso de que no encuentre una tabla.
                {
                    stillCount = false; //Ponemos la variable de control en falso para detener la b�squeda (y por tanto la cuenta).
                }
            }
            while (stillCount == true); //Indicamos que siga buscando y contando mientras la variable de control sea verdadera.

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Para hacer la gr�fica de ganancias.
            Excel.Chart[] Gr�ficasFrecuenciaFase = new Excel.Chart[numberOfTables];

            //Para trabajar con cada gr�fica.
            for (int i = 1; i <= numberOfTables; i++)
            {
                Gr�ficasFrecuenciaFase[i - 1] = currentWorkbook.Charts.Add(Type.Missing, Type.Missing, Type.Missing, Type.Missing); //Colocamos una hoja de trabajo para gr�ficas y la asignamos al objeto.
                               
                Gr�ficasFrecuenciaFase[i - 1].ChartWizard(
                    Gallery: XlChartType.xlLineMarkers,
                    PlotBy: XlRowCol.xlColumns,
                    CategoryLabels: 45,
                    SeriesLabels: 5,
                    HasLegend: true,
                    Title: allTheCells.Item[1 + ((i - 1) * 54), 1].Value,
                    CategoryTitle: "Frecuencias (Hz)",
                    ValueTitle: "Ganancias (dB)",
                    ExtraTitle: "Extra"
                    );

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
                Excel.SeriesCollection[] Colecci�nDeSeries = new Excel.SeriesCollection[numberOfTables]; //Hacemos un array de colecciones de series, en el que cada elemento...
                //es la colecci�n de series de una gr�fica en particular.
                Excel.Series[] SeriesDeDatos = new Excel.Series[5 * numberOfTables];
                Colecci�nDeSeries[i - 1] = Gr�ficasFrecuenciaFase[i - 1].SeriesCollection(); //Obtenenmos la colecci�n de series de la gr�fica.

                while (Colecci�nDeSeries[i - 1].Count > 0) //Mientras haya series en la colecci�n de series de una gr�fica.
                {
                    Colecci�nDeSeries[i - 1].Item(1).Delete();  //Borramos cada series.
                                                                //Cada vez que se borra una series, la siguiete serie se vuelve la primera.
                                                                //Borramos todas las series porque la gr�fica debe estar vac�a desde su creaci�n.

                }


                for (int j = 5 * (i - 1); j < 5 * (i); j++) //Ciclamos en las 5 mediciones.
                {
                    SeriesDeDatos[j] = Colecci�nDeSeries[i - 1].NewSeries(); //Creamos una nueva serie en la gr�fica.
                    //SeriesDeDatos[j].ClearFormats();
                    SeriesDeDatos[j].XValues = rangeOfFrequencies.Cells; //Introducimos el rango de celdas donde se encuentran las frecuencias.
                    SeriesDeDatos[j].Values = rangesOfPhase[j % 5]; //Introducimos el rango de celdas donde se encuentran las ganancias.
                    //Usamos modulo para sacar valores desde cero a 5.
                    SeriesDeDatos[j].Name = "Medici�n #" + ((j % 5) + 1).ToString(); //Le ponemos nombre a cada serie.
                }

                //Modificamos los ejes.
                Excel.Axes AxesOfTheChart = null; //Creamos un objeto de colecci�n de ejes.
                AxesOfTheChart = Gr�ficasFrecuenciaFase[i - 1].Axes(); //Enlazamos el objeto con los ejes de la gr�fica.
                Excel.Axis XAxis = AxesOfTheChart.Item(XlAxisType.xlCategory, XlAxisGroup.xlPrimary); //Obtenenmos el eje "X" primario.
                XAxis.HasMajorGridlines = true; //Activamos las l�neas de cuadr�cula mayores.
                XAxis.HasMinorGridlines = true; //Activamos las l�neas de cuadr�cula menores.
                XAxis.AxisBetweenCategories = false;

                //Hacemos los mismo para el eje "Y".
                Excel.Axis YAxis = AxesOfTheChart.Item(XlAxisType.xlValue, XlAxisGroup.xlPrimary); //Debemos sacar los primarios, porque son los que se pueden modificar.
                YAxis.HasMajorGridlines = true; //Activamos l�neas de cuadr�cula mayores.
                YAxis.HasMinorGridlines = true; //Activamos l�neas de cuadr�cula menores.

                //Modificamos las l�neas de cuadr�cula del eje "X".
                Excel.Gridlines L�neasMenoresEjeX = null; //Creamos el objeto interfaz.
                L�neasMenoresEjeX = XAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMenoresEjeX = L�neasMenoresEjeX.Border; //Creamos un objeto borde y lo enlazamos.
                BordeL�neasMenoresEjeX.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un n�mero hexadecimal.

                Excel.Gridlines L�neasMayoresEjeX = null; //Creamos el objeto interfaz.
                L�neasMayoresEjeX = XAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMayoresEjeX = L�neasMayoresEjeX.Border; //Creamos el objeto borde y lo enlazamos.
                BordeL�neasMayoresEjeX.Color = 0xD9D9D9; //Modificamos su color.

                //Modificamos las l�neas de cuadr�cula del eje "Y".
                Excel.Gridlines L�neasMenoresEjeY = null; //Creamos el objeto interfaz.
                L�neasMenoresEjeY = YAxis.MinorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMenoresEjeY = L�neasMenoresEjeY.Border; //Creamos un objeto borde y lo enlazamos.
                BordeL�neasMenoresEjeY.Color = 0xF2F2F2; //Modificamos su color; NOTA: Debe ser un n�mero hexadecimal.

                Excel.Gridlines L�neasMayoresEjeY = null; //Creamos el objeto interfaz.
                L�neasMayoresEjeY = YAxis.MajorGridlines; //Lo enlazamos.
                Excel.Border BordeL�neasMayoresEjeY = L�neasMayoresEjeY.Border; //Creamos el objeto borde y lo enlazamos.
                BordeL�neasMayoresEjeY.Color = 0xD9D9D9; //Modificamos su color.

                //MessageBox.Show(String.Format("Chart {0}", currentWorkbook.Charts.Item[currentWorkbook.Charts.Count].Name));
            }


        }

        private void BorraGr�ficas_Click(object sender, EventArgs e)
        {
            //Declaramos objetos que van a guardar los datos actuales del archivo de Excel.
            Excel._Application currentApplication = null;
            Excel.Workbooks currentWorkbooks = null;
            Excel._Workbook currentWorkbook = null;
            Excel.Sheets currentWorksheets = null;
            Excel.Worksheet currentWorksheet = null;
            Excel.Range allTheCells = null;
                       
            //Variable para indicar que el sistema fall� en encontrar el archivo de Excel abierto.
            bool wasFoundRunning = false;

            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            //////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


            //Checamos si Excel est� abierto
            try
            {
                currentApplication = (Excel._Application)GetOpenedExcelApplication(); //Guardamos la aplicaci�n de Excel en currentApplication.
                wasFoundRunning = true; //Lo ponemos en true en caso de que s� est� abierto.
            }
            catch (Exception ex) //Excel no est� abierto.
            {
                wasFoundRunning = false; //Lo ponemos en false en caso de que no est� abierto.
                MessageBox.Show("Failed to get opened Excel file", "Error: " + ex.ToString(), MessageBoxButtons.OK); //Indicamos fracaso
            }
            finally
            {
                if (currentApplication != null && wasFoundRunning == true) //Si no hubo excepciones.
                {
                    MessageBox.Show("Found Excel opened file", "Success"); //Indicamos �xito.
                }

            }

            currentWorkbooks = currentApplication.Workbooks; //Obtenemos la colecci�n de los libros de trabajo abiertos de la aplicaci�n Excel.
            currentWorkbook = currentWorkbooks.Item[1]; //Obtenemos el primer libro de trabajo abierto.
            currentWorksheets = currentWorkbook.Worksheets; //Obtenemos la �nica hoja de trabajo del libro de trabajo.
            currentWorksheet = (Excel.Worksheet)currentWorksheets.Item[1];
            allTheCells = currentWorksheet.Cells;

            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
            /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

            //Iteramos por cada hoja de trabajo que no sea la principal y las borramos.
            while (currentWorksheets.Count > 1)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.

                currentWorksheets.Item[1].Delete(); //Borramos las dem�s hojas de trabajo con datos.
                                                    //S�lo debe de haber una hoja de trabajo.
            }

            //Iteramos por cada hoja de gr�fica y las borramos.
            while (currentApplication.Charts.Count > 0)
            {
                currentApplication.WindowState = Excel.XlWindowState.xlMaximized; //Maximizamos la ventana de la aplicaci�n de Excel abierta.
                SetForegroundWindow(currentApplication.Hwnd); //Hacemos la ventana de la aplicaci�n est� al frente.

                currentWorkbook.Charts.Item[1].Delete(); //Borramos cada hoja de gr�fica.

            }
        }
    }

    //Creamos una clase con c�digo antiguo de Marshal que es necesario para este programa.
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



