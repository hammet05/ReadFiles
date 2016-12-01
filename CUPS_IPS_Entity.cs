using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;

namespace DAL
{
    public class CUPS_IPS_Entity
    {

        static public void BulkCopyFiles(DataTable dt)
        {
            try
            {
                var conexionBulkCopy = new SqlConnection { ConnectionString = DAL.Conexion.GetConection() };

                conexionBulkCopy.Open();
             
                var bc = new SqlBulkCopy(conexionBulkCopy);
           
                bc.DestinationTableName = "CUPS_IPS";              
                bc.WriteToServer(dt);
               
                conexionBulkCopy.Close();
            }
            catch (SqlException sqlException)
            {
                var m_Sb = new StringBuilder();
                var logError = ConfigurationManager.AppSettings["RutaLogError"];
                if (!File.Exists(logError))
                {
                    var line = m_Sb.ToString() + " " + "Error al guardar en base de datos:" + dt.TableName + " " + dt.Rows[0] + " " + sqlException.Message + Environment.NewLine;
                    File.WriteAllText(logError, line);
                }
                else
                {
                    FileInfo fi = new FileInfo(logError);
                    File.AppendAllText(fi.FullName, m_Sb.ToString() + "Error al guardar en base de datos:" + dt.TableName + " " + "Nombre Cups Error" + dt.Rows[0].ItemArray[2].ToString() + " " + sqlException.Message + Environment.NewLine);
                }
            }
            catch (Exception ex)
            {

                Console.WriteLine(ex.Message);
                Console.Read();
            }

        }
    }
}
