using DocumentFormat.OpenXml.Spreadsheet;
using MySql.Data.MySqlClient;

namespace Excel.Models.DAO
{
    public class DAOConexion
    {
        private static DAOConexion instancia;

        public static DAOConexion Instancia
        {
            get
            {
                if (instancia == null) instancia = new DAOConexion();
                return instancia;
            }
        }

 

        public string getConexionReporteador()
        {

            var strCon = "Server = localhost; User = root; Database = reporteador; port = 3306; Password = AxlPKT1.com";
            return strCon;
        }
    }
}
