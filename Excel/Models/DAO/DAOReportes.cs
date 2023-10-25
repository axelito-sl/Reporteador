using Excel.Database;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Org.BouncyCastle.Asn1.IsisMtt.X509;

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Threading;
using System.Web;
using System.Web.Http;
using Excel.Models;
using Excel.Models.Responses;

namespace Excel.Models.DAO
{
    public class DAOReportes
    {

        #region cafe
        /*
         
         public class café:bebida
        {
        float Mlagua=300;
        int cucharadas_cafe=2;
        bool hascream=false;
        bool lleno=false;
        public café()
        {
        
        this.checkbeber();
        
        }

        private void checkbeber()
        {
        if(this.lleno==false)
            this.beber();
        else()
        {
            this.refill();
        }
        }
        }
         
         */
        #endregion


        private static DAOReportes instance;

        public static DAOReportes Instance()
        {
            if (instance == null) instance = new DAOReportes();
            return instance;
        }

        MySqlConnection objConn = new MySqlConnection(DAOConexion.Instancia.getConexionReporteador());




        public dynamic getreportes() {
            objConn = new MySqlConnection(DAOConexion.Instancia.getConexionReporteador());
            var query = "select * from master_reportes";
            MySqlCommand cmd = new MySqlCommand(query, objConn);

            STDResponse response = new STDResponse()
            {
                isError = false,
                Message = "Proceso Exitoso",
                totalRecords = 0,
                data = new List<object>()
            };
            try
            {

                objConn.Open();
                var registros = cmd.ExecuteReader();
                while (registros.Read())
                {
                    master_reportes report = new master_reportes();
                    report.id = int.Parse(registros["id"].ToString());
                    report.nombre = registros["nombre"].ToString();
                    report.descripcion = registros["descripcion"].ToString();
                    report.num_columnas = int.Parse(registros["num_columnas"].ToString());
                    report.Ordenamiento = registros["Ordenamiento"].ToString();
                    report.Filtros = registros["Filtros"].ToString();
                    report.estatus = registros["estatus"].ToString();
                    response.data.Add(report);
                    response.totalRecords++;
                }



            }
            catch (Exception ex)
            {
                return (ex.Message.ToString());
            }
            finally { objConn.Close(); }


            return response;
        }

        public string savereporte(int cliente,int usuario,string nombre,string desc,int columnas,string ordenamiento,string filtros, List<(string,string)>nomcolumnas)
        {
            objConn = new MySqlConnection(DAOConexion.Instancia.getConexionReporteador());
            var query = "INSERT INTO `reporteador`.`master_reportes` (`id_cliente`, `id_usuario_creador`, `nombre`, `descripcion`, `num_columnas`, `Ordenamiento`, `Filtros`, `estatus`) VALUES (@cliente,@usuario,@nombre,@desc,@columnas,@ordenamiento, @filtros, 'A');";
            MySqlCommand cmd = new MySqlCommand(query, objConn);
            cmd.Parameters.AddWithValue("@cliente",cliente);
            cmd.Parameters.AddWithValue("@usuario", usuario);
            cmd.Parameters.AddWithValue("@nombre", nombre);
            cmd.Parameters.AddWithValue("@desc",desc );
            cmd.Parameters.AddWithValue("@columnas",columnas);
            cmd.Parameters.AddWithValue("@ordenamiento",ordenamiento );
            cmd.Parameters.AddWithValue("@filtros",filtros );

            try
            {

                objConn.Open();
                var result = cmd.ExecuteNonQuery();



                //insertar las columnas del reporte recién generado en caso de que el reporte se haya guardado bien
               
               

               
                if (result != -1)
                {
                    

                    for (int i = 0; i < columnas; i++)
                    {
                        query = "INSERT INTO `reporteador`.`columnas_reportes` (`id_reporte`, `nom_columna`, `nombre_tipo_dato`, `operaciones_disp`, `estatus`) VALUES ((select id from master_reportes where id_cliente=@cliente order by id desc limit 1),@nomcol,@datocol, (select config_tipos_datos.operaciones_disp from config_tipos_datos where @datocol=config_tipos_datos.nombre_tipo_dato) ,'A');";
                        MySqlCommand cmd2 = new MySqlCommand(query, objConn);
                        cmd2.Parameters.AddWithValue("@cliente",cliente);
                        cmd2.Parameters.AddWithValue("@nomcol", nomcolumnas[i].Item1);
                        cmd2.Parameters.AddWithValue("@datocol", nomcolumnas[i].Item2);
                        var result2=cmd2.ExecuteNonQuery();
                        if (result2 != -1)
                            continue;
                        break;
                    }
                }




                return "ok";
            }catch(Exception ex)
            {
                return ex.Message.ToString();
            }
            finally
            {

                objConn.Close();
            }







        }


      

}
}