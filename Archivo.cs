using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;

namespace LectorArchivos
{
    public class Archivo
    {
        public string Path { get; set; }
        public string Folder { get; set; }
        public string Sheet { get; set; }
       
        public static List<Archivo> ListaArchivosProcesar { get; set; }

        static bool Procesado = false;

        public Archivo() { }
        public Archivo(string path, string folder, string sheet)
        {
            Path = path;
            Folder = folder;
            Sheet = sheet;
        }
        public Archivo(string path, string sheet)
        {
            Path = path;
            Sheet = sheet;
        }

        public static List<Archivo> ObtenerRutasArchivosProcesar()
        {
            //var path = @"E:\Documentacion\Mapas\Requerimientos\Tarifas\Archivos Prueba\SANDRA M BENAVIDES";
            //var ruta = string.Empty;
            //var lArchivosExcel = new List<string>(Directory.GetFiles(path, "PAR*.*", SearchOption.AllDirectories));
            //var xlAPP = new Microsoft.Office.Interop.Excel.Application();

            //var lHojasExcel = new List<Archivo>();
            //string hoja = string.Empty;
            //ListaArchivosProcesar = new List<Archivo>();
            //var taskProceso = new Task(ProcesarArchivos);
            //taskProceso.Start();

            //taskProceso.Wait();
            //try
            //{
            //    var inicio = DateTime.Now;

            //    foreach (var item in lArchivosExcel)
            //    {
            //        var xlWbk = xlAPP.Workbooks.Open(item);

            //        Procesado = false;
            //        foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWbk.Worksheets)
            //        {
            //            var archivoExcel = new Archivo();
            //            archivoExcel.Path = item;
            //            archivoExcel.Sheet = sheet.Name;

            //            try
            //            {
            //                if (item.Contains(sheet.Name))
            //                {
            //                    lHojasExcel.Add(archivoExcel);
            //                    ListaArchivosProcesar.Add(archivoExcel);
            //                    Procesado = true;
            //                    xlWbk.Close();
            //                    break;
            //                }


            //            }
            //            catch (IOException ioex)
            //            {

            //                Console.WriteLine(ioex.Message);
            //            }
            //            catch (Exception ex)
            //            {
            //                Console.WriteLine(ex.Message);
            //            }

            //        }

            //        if (Procesado == false)
            //        {
            //            var m_Sb = new StringBuilder();
            //            var logError = ConfigurationManager.AppSettings["RutaLogError"];
            //            if (!File.Exists(logError))
            //            {
            //                var line = m_Sb.ToString() + " " + "Carpeta:" + item + Environment.NewLine;
            //                File.WriteAllText(logError, line);
            //            }
            //            else
            //            {
            //                FileInfo fi = new FileInfo(logError);
            //                File.AppendAllText(fi.FullName, m_Sb.ToString() + " " + "Carpeta:" + item + Environment.NewLine);
            //            }
            //            xlWbk.Close();
            //        }



            //    }
            //    var fin = DateTime.Now;
            //    TimeSpan diff = fin - inicio;

            //    Console.WriteLine("Inicio= {0} - Fin= {1} - Diferencia= {2}:{3}:{4}.{5},{6}",
            //                      inicio.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
            //                      fin.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
            //                      diff.Hours, diff.Minutes, diff.Seconds, diff.Milliseconds,
            //                      diff.Milliseconds);
            //    Console.Read();
            //}
            //catch (Exception ex)
            //{

            //    Console.WriteLine(ex.Message);
            //}
     
            
            ////xlAPP.Quit();

            ListaArchivosProcesar = new List<Archivo>();
             var t= ProcesarArchivos();
            //taskProceso.Start();
             t.Wait();
            return ListaArchivosProcesar;
        }


        static async Task<List<Archivo>> ProcesarArchivos()
        {
            await Task.Delay(1);
            var path = @"E:\Documentacion\Mapas\Requerimientos\Tarifas\Archivos Prueba\ANGELA MORENO";
            var ruta = string.Empty;
            var lArchivosExcel = new List<string>();
            
            if (path.Contains("ANGELA"))
            {
                lArchivosExcel = Directory.GetFiles(path, "*.xls*", SearchOption.AllDirectories).ToList();
            }
            else
            {
                lArchivosExcel = Directory.GetFiles(path, "PAR*.xls*", SearchOption.AllDirectories).ToList();
            }


            var xlAPP = new Microsoft.Office.Interop.Excel.Application();

            var lHojasExcel = new List<Archivo>();
            string hoja = string.Empty;
            ListaArchivosProcesar = new List<Archivo>();
            try
            {
                var inicio = DateTime.Now;

                foreach (var item in lArchivosExcel)
                {
                    var xlWbk = xlAPP.Workbooks.Open(item);

                    Procesado = false;
                    foreach (Microsoft.Office.Interop.Excel.Worksheet sheet in xlWbk.Worksheets)
                    {
                        var archivoExcel = new Archivo();
                        archivoExcel.Path = item;
                        archivoExcel.Sheet = sheet.Name;

                        try
                        {
                            if (item.Contains(sheet.Name)|| sheet.Name.StartsWith("TARIFAS"))
                            {
                                lHojasExcel.Add(archivoExcel);
                                ListaArchivosProcesar.Add(archivoExcel);
                                Procesado = true;
                                xlWbk.Close();
                                break;
                            }


                        }
                        catch (IOException ioex)
                        {

                            Console.WriteLine(ioex.Message);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }

                    }

                    if (Procesado == false)
                    {
                        var m_Sb = new StringBuilder();
                        var logError = ConfigurationManager.AppSettings["RutaLogError"];
                        if (!File.Exists(logError))
                        {
                            var line = m_Sb.ToString() + " " + "Carpeta:" + item + Environment.NewLine;
                            File.WriteAllText(logError, line);
                        }
                        else
                        {
                            FileInfo fi = new FileInfo(logError);
                            File.AppendAllText(fi.FullName, m_Sb.ToString() + " " + "Carpeta:" + item + Environment.NewLine);
                        }
                        xlWbk.Close();
                    }



                }
                var fin = DateTime.Now;
                TimeSpan diff = fin - inicio;

                Console.WriteLine("Inicio= {0} - Fin= {1} - Diferencia= {2}:{3}:{4}.{5},{6}",
                                  inicio.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
                                  fin.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
                                  diff.Hours, diff.Minutes, diff.Seconds, diff.Milliseconds,
                                  diff.Milliseconds);
               
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }


            //xlAPP.Quit();
            return ListaArchivosProcesar;
        }
    }
}
