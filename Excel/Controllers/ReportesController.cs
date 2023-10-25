using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Math;
using Excel.Database;
using Excel.Models;
using Excel.Models.DAO;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Hosting;
using MySql.Data.MySqlClient;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Data.Entity;
using System.Text;
namespace Excel.Controllers
{
    [ApiController]
    [Route("Api/[controller]")]
    public class ReportesController : Controller
    {


  
        public ReportesController()
        {
           
        }


        [HttpGet]
        [Route("GetReportes")]

        public object getreportes()
        {
           
            var reportes=DAOReportes.Instance().getreportes();
           return reportes;
        }




        [HttpPost]
        [Route("AddReporte")]

        public IActionResult SaveReporteDinamico([FromBody] object data)
        {

            var Jdata = JsonConvert.DeserializeObject<dynamic>(data.ToString());
            string filtro = JsonConvert.SerializeObject(Jdata.filtro);
            string orden = JsonConvert.SerializeObject(Jdata.orden);

          
            string nc = Jdata.nomcolumnas.ToString();

            // Deserializa el JSON en un objeto dinámico
            dynamic jsonObject = JsonConvert.DeserializeObject(Jdata.nomcolumnas.ToString());

            // Accede al arreglo "nomcolumnas" y conviértelo en una lista de tuplas
            List<(string, string)> lista = new List<(string, string)>();
            foreach (var item in jsonObject)
            {
                lista.Add(((string)item[0], (string)item[1]));
            }

            var a = DAOReportes.Instance().savereporte((int)Jdata.cliente, (int)Jdata.usuario,Jdata.nombre.ToString(),Jdata.desc.ToString(), (int)Jdata.columnas,orden,filtro,lista);
           
                return Ok();
            
        }

        [HttpDelete]
        [Route("DeleteReporte")]
        public IActionResult DelReporte()
        {
            return Ok();
        }
       
            
             
            


        }


    }

