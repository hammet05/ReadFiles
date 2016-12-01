using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DAL
{
    public class Conexion
    {
        public static string GetConection()
        {
            string conex = null;
            using (var data = new ArchivosContextDataContext())
            {
                conex = data.Connection.ConnectionString;
            }
            return conex;
        }
      
    }
}
