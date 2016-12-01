using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Formatters.Binary;
using System.Data.SqlClient;
using System.Data;
using System.Data.OleDb;
using System.ComponentModel;
using System.Configuration;
using DAL;

namespace LectorArchivos
{
    class Program
    {
       
        static void Main(string[] args)
        {

            try
            {
               
                Task task = new Task(ProcesarDatos);
                task.Start();

                task.Wait();
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.Read();
            }
           
                     

            //TimeSpan diff = fin - inicio;

            //Console.WriteLine("Inicio= {0} - Fin= {1} - Diferencia= {2}:{3}:{4}.{5},{6}",
            //                  inicio.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
            //                  fin.ToString("yyyy/MM/dd HH:mm:ss.fffffff"),
            //                  diff.Hours, diff.Minutes, diff.Seconds, diff.Milliseconds,
            //                  diff.Milliseconds);
            //Console.Read();
        }

       
        static async void ProcesarDatos()
        {
           // await Task.Delay(100);
          

            var ruta = string.Empty;
            var hoja = string.Empty;
            var listaArchivosProcesar = Archivo.ObtenerRutasArchivosProcesar();
            foreach (var item in listaArchivosProcesar)
            {
             
                var task = LeerArchivo(item.Path, item.Sheet);

                var task2 = ConvertirListaToDataTable(task.Result);

                var task3 = GuardarArchivo(task2.Result);
            }


              
           //}
        
          
           
        }
        

        static async Task<List<CUPS_IPS>>LeerArchivo(string archivo, string hoja)
        {
            await Task.Delay(1000); 
            var listaCupsIPS = new List<CUPS_IPS>();

            try
            {

                string sql = null;
                var dt = new DataTable();
                var conex = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + archivo + ";Extended Properties=Excel 12.0;");
                var cmd = conex.CreateCommand();
                cmd.Connection = conex;
                cmd.CommandType = CommandType.Text;             
                sql = "Select * from [" + hoja + "$b1:C10000] ";               
               
                cmd.CommandText = sql;
                var adapter = new OleDbDataAdapter(sql, conex);
                adapter.Fill(dt);

                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow item in dt.Rows)
                    {
                        var cupsIPS = new CUPS_IPS()
                        {
                            Codigo = string.IsNullOrEmpty(item[0].ToString()) ? "0" : item[0].ToString(),
                            Nombre = string.IsNullOrEmpty(item[1].ToString()) ? "SIN DATOS" : item[1].ToString()
                        };
                        var cupsAgregarLista = listaCupsIPS.Find(x => x.Nombre == cupsIPS.Nombre);
                        if (cupsAgregarLista == null)
                        {
                            listaCupsIPS.Add(cupsIPS);
                        }
                    }

                }
            }
           
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.Read();
            }
            
            return listaCupsIPS;
        }

        static async Task<DataTable> ConvertirListaToDataTable(List<CUPS_IPS>lCupsIPS)
        {
            await Task.Delay(1000); 
            var nuevoDT = CrearDatatable();
            try
            {
              
                foreach (var cups in lCupsIPS)
                {
                    var newFila = nuevoDT.NewRow();
                    newFila["CODIGO"] = cups.Codigo;
                    newFila["NOMBRE"] = cups.Nombre;
                    newFila["ESTADO"] = true;

                    nuevoDT.Rows.Add(newFila);
                }
                
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
            }

            return nuevoDT;
            
        }
        public static DataTable CrearDatatable()
        {
            var dtCupsIPS = new DataTable("xxx");
            var Id = new DataColumn("ID")
            {
                DataType = typeof(int),
                ColumnMapping = MappingType.Attribute,
                AutoIncrement=true,
                AutoIncrementSeed=1,
                AutoIncrementStep=1
            };
            var codigoCUPS = new DataColumn("CODIGO")
            {
                DataType = typeof(string),
                ColumnMapping = MappingType.Attribute
            };

            var nombreCUPS = new DataColumn("NOMBRE")
            {
                DataType = typeof(string),
                ColumnMapping = MappingType.Attribute
            };

            var estado = new DataColumn("ESTADO")
            {
                DataType = typeof(bool),
                ColumnMapping = MappingType.Attribute
            };
            dtCupsIPS.Columns.Add(Id);
            dtCupsIPS.Columns.Add(codigoCUPS);
            dtCupsIPS.Columns.Add(nombreCUPS);
            dtCupsIPS.Columns.Add(estado);

            return dtCupsIPS;
        }

        static async Task GuardarArchivo(DataTable dt)
        {
            //await Task.Delay(100); 
            CUPS_IPS_Entity.BulkCopyFiles(dt);
        }

      
    }

}
