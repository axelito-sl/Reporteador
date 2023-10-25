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
using System.Web.Http.Results;

namespace Excel.Models.DAO
{
public class DAOPlanes
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


        private static DAOPlanes instance;

        public static DAOPlanes Instance()
        {
            if (instance == null) instance = new DAOPlanes();
            return instance;
        }

        MySqlConnection objConn = new MySqlConnection(DAOConexion.Instancia.getConexionReporteador());




    public dynamic getplanes(){
            objConn = new MySqlConnection(DAOConexion.Instancia.getConexionReporteador());
            var query = "select * from planes";
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
                    Planes report = new Planes();
                    report.id = int.Parse(registros["id"].ToString());
                    report.nombre = registros["nombre"].ToString();
                    report.estatus = registros["estatus"].ToString();
                    response.data.Add(report);
                    response.totalRecords++;
                }

                
            
            }
            catch(Exception ex)
            {
                return (ex.Message.ToString());
            }
            finally { objConn.Close(); }
            
            
            return response;
    }



        public object delreporte()
        {
            return OkNegotiatedContentResult()
        }





}
}