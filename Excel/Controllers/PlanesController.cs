using Excel.Models.DAO;
using Microsoft.AspNetCore.Mvc;

namespace Excel.Controllers
{
    [ApiController]
    [Route("Api/[controller]")]
    public class PlanesController : Controller
    {

        public PlanesController()
        {
            
        }


        [HttpGet]
        [Route("GetPlanes")]

        public object getplanes()
        {
           
            var planes=DAOPlanes.Instance().getplanes();
           return planes;
        }
       

        // [HttpGet]
        // [Route("GetPlanes")]

        

        // [HttpPost]
        // [Route("GuardarReporte")]

        // public ActionResult SaveReporteDinamico([FromBody] object data) 
        // {

        //     try
        //     {

        //         var Jdata = JsonConvert.DeserializeObject<dynamic>(data.ToString());
        //         //insertar un nuevo reporte en la tabla master_reportes




        //         master_reportes reporte = new master_reportes()
        //         {
        //             nombre = Jdata.Nombre.ToString(),
        //             descripcion = Jdata.Descripcion.ToString(),
        //             num_columnas = Jdata.Num_columnas,
        //             Ordenamiento = Jdata.Ordenamiento.ToString(),
        //             Filtros=Jdata.Filtros.ToString(),
        //             estatus = "A"
        //         };

        //         contexto.master_reportes.Add(reporte);
        //         contexto.SaveChanges();


        //         //insertamos las columnas en columnas_reportes

        //         for(int i=0;i<Convert.ToInt32(Jdata.Num_columnas);i++){
                    
        //         }


        //         return Ok();





        //     }
        //     catch(Exception e)
        //     {
        //         return BadRequest(e.Message);
        //     }


        // }


    }
}
