using Excel.Database;
using Microsoft.AspNetCore.Mvc;
using SpreadsheetLight;
using SpreadsheetLight.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

using System.Net;
using System.Net.Http.Headers;
using System.Net.Mime;
using DocumentFormat.OpenXml.Bibliography;
using GroupDocs.Conversion.Options.Convert;
using GroupDocs.Conversion.Options.Load;
using GroupDocs.Watermark.Search;
using GroupDocs.Watermark;

namespace Excel.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController: ControllerBase
    {
        private readonly ConexionContext contexto;
        private readonly IWebHostEnvironment _hostEnvironment;
       

        public ExcelController(ConexionContext contexto,IWebHostEnvironment hostEnvironment)
        {
            this.contexto = contexto;
            _hostEnvironment = hostEnvironment;
           
        }





        #region a
        //[HttpGet]
        //[Route("GetExcel")]

        //public ActionResult<dynamic> Excel()
        //{
        //    var registros = contexto.pruebaexcel.ToList();


        //    SLDocument sl = new SLDocument();
        //    int cabeceraActual = 2;
        //    int cabeceraInicial = 2;
        //    int contRegistros = 0;
        //    sl.MergeWorksheetCells("A1", "B1");
        //    sl.MergeWorksheetCells("C1", "E1");
        //    sl.MergeWorksheetCells("F1", "H1");
        //    sl.MergeWorksheetCells("I1", "J1");
        //    //bitmap para la foto
        //    System.Drawing.Bitmap bm = new System.Drawing.Bitmap("C:\\Users\\PKT1-Denisse\\source\\repos\\Excel\\Excel\\pkt.jpg");

        //    //estilo para encabezados de reporte
        //    SLStyle encabezados = sl.CreateStyle();
        //    encabezados.Font.Bold = true;
        //    encabezados.SetVerticalAlignment(VerticalAlignmentValues.Center);//alinear al medio
        //    encabezados.Alignment.WrapText = true;
        //    //valores de encabezado reporte
        //    string reportenombre = "REPORTE DE SALDO POR FRANQUICIA";
        //    string franquicia = "Franquicia: PRO";
        //    string fechagen = "Generado el " + DateTime.Now.ToString();

        //    sl.SetCellValue("C1", reportenombre);
        //    sl.SetCellValue("F1", franquicia);
        //    sl.SetCellValue("I1", fechagen);
        //    sl.SetCellStyle("A1", "I1", encabezados);


        //    byte[] ba;
        //    using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
        //    {
        //        bm.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
        //        ms.Close();
        //        ba = ms.ToArray();

        //    }
        //    SLPicture pic = new SLPicture(ba, DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg);
        //    pic.SetPosition(0, 0);
        //    pic.ResizeInPixels(180, 70);
        //    sl.InsertPicture(pic);
        //    sl.SetRowHeight(1, 55);//ancho de fila
        //    //DEFINIMOS LOS ESTILOS PARA LAS COLUMNAS 

        //    SLStyle CabStyle = sl.CreateStyle();
        //    CabStyle.Font.FontName = "Calibri";
        //    CabStyle.Font.FontSize = 11;
        //    CabStyle.Font.Bold = true;
        //    CabStyle.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, System.Drawing.Color.ForestGreen, System.Drawing.Color.DarkOliveGreen);
        //    sl.SetCellStyle("A" + cabeceraInicial, "J" + cabeceraInicial, CabStyle);

        //    //estilo para bordes de celdas
        //    SLStyle borders = sl.CreateStyle();
        //    borders.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
        //    borders.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
        //    borders.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
        //    borders.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;


        //    //estilo para monedas
        //    SLStyle currency = sl.CreateStyle();
        //    currency.FormatCode = "$,###,###,###.00";


        //    //ESCRIBIMOS LOS ENCABEZADOS DE LAS COLUMNAS DEL REPORTE
        //    sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "SALDO POR FRANQUICIA");
        //    sl.SetCellValue("A" + cabeceraInicial, "ID CLIENTE");
        //    sl.SetCellValue("B" + cabeceraInicial, "NOMBRE CLIENTE");
        //    sl.SetCellValue("C" + cabeceraInicial, "SALDO INICIAL");
        //    sl.SetCellValue("D" + cabeceraInicial, "ENERO");
        //    sl.SetCellValue("E" + cabeceraInicial, "FEBRERO");
        //    sl.SetCellValue("F" + cabeceraInicial, "MARZO");
        //    sl.SetCellValue("G" + cabeceraInicial, "ABRIL");
        //    sl.SetCellValue("H" + cabeceraInicial, "MAYO");
        //    sl.SetCellValue("I" + cabeceraInicial, "JUNIO");
        //    sl.SetCellValue("J" + cabeceraInicial, "TOTAL POR CLIENTE");

        //    //RELLENAMOS CON LOS REGISTROS
        //    foreach (var item in registros)
        //    {
        //        contRegistros++;
        //        cabeceraActual++;
        //        sl.SetCellValue("A" + cabeceraActual, item.id_cliente);
        //        sl.SetCellValue("B" + cabeceraActual, item.nombre_cliente.ToString());
        //        sl.SetCellValue("C" + cabeceraActual, item.saldo_inicial);
        //        sl.SetCellValue("D" + cabeceraActual, item.enero);
        //        sl.SetCellValue("E" + cabeceraActual, item.febrero);
        //        sl.SetCellValue("F" + cabeceraActual, item.marzo);
        //        sl.SetCellValue("G" + cabeceraActual, item.abril);
        //        sl.SetCellValue("H" + cabeceraActual, item.mayo);
        //        sl.SetCellValue("I" + cabeceraActual, item.junio);
        //        //    sl.SetCellValue("J" + cabeceraActual,"=SUM(C"+"{cabeceraActual}:I{cabeceraActual}");
        //        //sl.SetCellValue("J" + cabeceraActual, "=SUM(C" + $"{cabeceraActual}:I"+$"{cabeceraActual}");
        //        //  sl.SetCellValue("J" + cabeceraActual, "=SUMA(C" + cabeceraActual + ":I" + cabeceraActual + ")");
        //        sl.SetCellValue("J" + cabeceraActual, "=SUM(C" + cabeceraActual + ":I" + cabeceraActual + ")");

        //    };


        //    sl.SetCellStyle("C" + cabeceraInicial, "J" + cabeceraActual, currency);

        //    sl.SetCellStyle("C" + (contRegistros + 3), "J" + (contRegistros + 3), currency);

        //    //estilo para columna de total_cliente
        //    SLStyle totales = sl.CreateStyle();
        //    totales.Font.FontName = "Calibri";
        //    totales.Font.FontSize = 11;
        //    totales.Font.Bold = true;

        //    sl.SetCellStyle("J" + cabeceraInicial, "J" + cabeceraActual, totales);

        //    sl.SetCellStyle("A" + cabeceraInicial, "J" + cabeceraActual, borders);

        //    //suma de columnas
        //    //for(int i= 0; i < contRegistros - 2; i++)
        //    //{
        //    //    sl.SetCellValue("C" + (i + cabeceraInicial), "=SUM(C:" + ")");
        //    //}
        //    sl.SetCellValue("C" + (contRegistros + 3), "=SUM(C" + (1 + cabeceraInicial) + ":C" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("D" + (contRegistros + 3), "=SUM(D" + (1 + cabeceraInicial) + ":D" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("E" + (contRegistros + 3), "=SUM(E" + (1 + cabeceraInicial) + ":E" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("F" + (contRegistros + 3), "=SUM(F" + (1 + cabeceraInicial) + ":F" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("G" + (contRegistros + 3), "=SUM(G" + (1 + cabeceraInicial) + ":G" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("H" + (contRegistros + 3), "=SUM(H" + (1 + cabeceraInicial) + ":H" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("I" + (contRegistros + 3), "=SUM(I" + (1 + cabeceraInicial) + ":I" + (contRegistros + 2) + ")");
        //    sl.SetCellValue("J" + (contRegistros + 3), "=SUM(J" + (1 + cabeceraInicial) + ":J" + (contRegistros + 2) + ")");
        //    //ajusta los tamaños de columna de acuerdo a el contenido que tienen
        //    sl.AutoFitColumn("A","L",20);


        //    sl.SaveAs("PRUEBA.xlsx");
        //    return Ok();
        //}
        #endregion;

        [HttpGet]
        [Route("GenerarReporte/{tipo}")]
        public async Task <ActionResult<dynamic>> Descarga(string nombredoc,string tipo)
        {

            var ruta = _hostEnvironment.ContentRootPath;
            var registros = contexto.pruebaexcel.ToList();


            SLDocument sl = new SLDocument();
            int cabeceraActual = 2;
            int cabeceraInicial = 2;
            int contRegistros = 0;
            sl.MergeWorksheetCells("A1", "B1");
            sl.MergeWorksheetCells("C1", "E1");
            sl.MergeWorksheetCells("F1", "H1");
            sl.MergeWorksheetCells("I1", "J1");
            //bitmap para la foto
            System.Drawing.Bitmap bm = new System.Drawing.Bitmap(ruta+"pkt.jpg");

            //estilo para encabezados de reporte
            SLStyle encabezados = sl.CreateStyle();
            encabezados.Font.Bold = true;
            encabezados.SetVerticalAlignment(VerticalAlignmentValues.Center);//alinear al medio
            encabezados.Alignment.WrapText = true;
            //valores de encabezado reporte
            string reportenombre = "REPORTE DE SALDO POR FRANQUICIA";
            string franquicia = "Franquicia: PRO";
            string fechagen = "Generado el " + DateTime.Now.ToString();

            sl.SetCellValue("C1", reportenombre);
            sl.SetCellValue("F1", franquicia);
            sl.SetCellValue("I1", fechagen);
            sl.SetCellStyle("A1", "I1", encabezados);


            byte[] ba;
            using (System.IO.MemoryStream ms = new System.IO.MemoryStream())
            {
                bm.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                ms.Close();
                ba = ms.ToArray();

            }
            SLPicture pic = new SLPicture(ba, DocumentFormat.OpenXml.Packaging.ImagePartType.Jpeg);
            pic.SetPosition(0, 0);
            pic.ResizeInPixels(180, 70);
            sl.InsertPicture(pic);
            sl.SetRowHeight(1, 55);//ancho de fila
            //DEFINIMOS LOS ESTILOS PARA LAS COLUMNAS 

            SLStyle CabStyle = sl.CreateStyle();
            CabStyle.Font.FontName = "Calibri";
            CabStyle.Font.FontSize = 11;
            CabStyle.Font.Bold = true;
            CabStyle.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, System.Drawing.Color.ForestGreen, System.Drawing.Color.DarkOliveGreen);
            sl.SetCellStyle("A" + cabeceraInicial, "J" + cabeceraInicial, CabStyle);

            //estilo para bordes de celdas
            SLStyle borders = sl.CreateStyle();
            borders.Border.LeftBorder.BorderStyle = BorderStyleValues.Thin;
            borders.Border.RightBorder.BorderStyle = BorderStyleValues.Thin;
            borders.Border.TopBorder.BorderStyle = BorderStyleValues.Thin;
            borders.Border.BottomBorder.BorderStyle = BorderStyleValues.Thin;


            //estilo para monedas
            SLStyle currency = sl.CreateStyle();
            currency.FormatCode = "$##,###,###.00";


            //ESCRIBIMOS LOS ENCABEZADOS DE LAS COLUMNAS DEL REPORTE
            sl.RenameWorksheet(SLDocument.DefaultFirstSheetName, "SALDO POR FRANQUICIA");
            sl.SetCellValue("A" + cabeceraInicial, "ID CLIENTE");
            sl.SetCellValue("B" + cabeceraInicial, "NOMBRE CLIENTE");
            sl.SetCellValue("C" + cabeceraInicial, "SALDO INICIAL");
            sl.SetCellValue("D" + cabeceraInicial, "ENERO");
            sl.SetCellValue("E" + cabeceraInicial, "FEBRERO");
            sl.SetCellValue("F" + cabeceraInicial, "MARZO");
            sl.SetCellValue("G" + cabeceraInicial, "ABRIL");
            sl.SetCellValue("H" + cabeceraInicial, "MAYO");
            sl.SetCellValue("I" + cabeceraInicial, "JUNIO");
            sl.SetCellValue("J" + cabeceraInicial, "TOTAL POR CLIENTE");

            //RELLENAMOS CON LOS REGISTROS
            foreach (var item in registros)
            {
                contRegistros++;
                cabeceraActual++;
                sl.SetCellValue("A" + cabeceraActual, item.id_cliente);
                sl.SetCellValue("B" + cabeceraActual, item.nombre_cliente.ToString());
                sl.SetCellValue("C" + cabeceraActual, item.saldo_inicial);
                sl.SetCellValue("D" + cabeceraActual, item.enero);
                sl.SetCellValue("E" + cabeceraActual, item.febrero);
                sl.SetCellValue("F" + cabeceraActual, item.marzo);
                sl.SetCellValue("G" + cabeceraActual, item.abril);
                sl.SetCellValue("H" + cabeceraActual, item.mayo);
                sl.SetCellValue("I" + cabeceraActual, item.junio);
                //    sl.SetCellValue("J" + cabeceraActual,"=SUM(C"+"{cabeceraActual}:I{cabeceraActual}");
                //sl.SetCellValue("J" + cabeceraActual, "=SUM(C" + $"{cabeceraActual}:I"+$"{cabeceraActual}");
                //  sl.SetCellValue("J" + cabeceraActual, "=SUMA(C" + cabeceraActual + ":I" + cabeceraActual + ")");
                sl.SetCellValue("J" + cabeceraActual, "=SUM(C" + cabeceraActual + ":I" + cabeceraActual + ")");

            };


            sl.SetCellStyle("C" + cabeceraInicial, "J" + cabeceraActual, currency);

            sl.SetCellStyle("C" + (contRegistros + 3), "J" + (contRegistros + 3), currency);

            sl.SetCellValue("B" + (contRegistros + 3), "TOTAL: ");

            //estilo para columna de total_cliente
            SLStyle totales = sl.CreateStyle();
            totales.Font.FontName = "Calibri";
            totales.Font.FontSize = 9;
            totales.Font.Bold = true;
            totales.FormatCode = "$##,###,###.00";
            totales.Fill.SetPattern(DocumentFormat.OpenXml.Spreadsheet.PatternValues.Solid, System.Drawing.Color.Gray, System.Drawing.Color.Gray);
            //   sl.SetCellStyle("A" + cabeceraInicial, "J" + cabeceraActual, totales);




            //suma de columnas
            //for(int i= 0; i < contRegistros - 2; i++)
            //{
            //    sl.SetCellValue("C" + (i + cabeceraInicial), "=SUM(C:" + ")");
            //}
            sl.SetCellValue("C" + (contRegistros + 3), "=SUM(C" + (1 + cabeceraInicial) + ":C" + (contRegistros + 2) + ")");
            sl.SetCellValue("D" + (contRegistros + 3), "=SUM(D" + (1 + cabeceraInicial) + ":D" + (contRegistros + 2) + ")");
            sl.SetCellValue("E" + (contRegistros + 3), "=SUM(E" + (1 + cabeceraInicial) + ":E" + (contRegistros + 2) + ")");
            sl.SetCellValue("F" + (contRegistros + 3), "=SUM(F" + (1 + cabeceraInicial) + ":F" + (contRegistros + 2) + ")");
            sl.SetCellValue("G" + (contRegistros + 3), "=SUM(G" + (1 + cabeceraInicial) + ":G" + (contRegistros + 2) + ")");
            sl.SetCellValue("H" + (contRegistros + 3), "=SUM(H" + (1 + cabeceraInicial) + ":H" + (contRegistros + 2) + ")");
            sl.SetCellValue("I" + (contRegistros + 3), "=SUM(I" + (1 + cabeceraInicial) + ":I" + (contRegistros + 2) + ")");
            sl.SetCellValue("J" + (contRegistros + 3), "=SUM(J" + (1 + cabeceraInicial) + ":J" + (contRegistros + 2) + ")");
            //ajusta los tamaños de columna de acuerdo a el contenido que tienen

            sl.SetCellStyle("B" + (contRegistros + 3), "J" + (contRegistros + 3), totales);
            sl.SetCellStyle("A" + cabeceraInicial, "J" + cabeceraActual, borders);

            sl.AutoFitColumn("A", "J");


            sl.SaveAs(nombredoc+".xlsx");
       
             

            string FilePath; // Ruta al archivo
            if (tipo == "pdf")
            {
                 FilePath = ruta + nombredoc + ".pdf"; // Ruta al archivo PDF
                Func<LoadOptions> loadOptions = () => new SpreadsheetLoadOptions
                {
                    OnePagePerSheet = true
                };
                using (var converter = new GroupDocs.Conversion.Converter(ruta + nombredoc + ".xlsx", loadOptions))
                {
                    // Convierta y guarde la hoja de cálculo en formato PDF
                    converter.Convert(ruta + nombredoc + ".pdf", new PdfConvertOptions());
                }

                // Configuramos el encabezado Content-Disposition para visualizar el archivo en el navegador
               Response.Headers["Content-Disposition"] = "inline; filename=" + nombredoc+".pdf";

                //// Configuramos el tipo de contenido
               Response.ContentType = "application/pdf";

                //// Leemos el archivo en bytes
               byte[] archivoBytes = System.IO.File.ReadAllBytes(FilePath);
                ////borro el archivo pdf generado del servidor
                System.IO.File.Delete(FilePath);
                //borramos el Excel generado del servidor
                System.IO.File.Delete(ruta+nombredoc+".xlsx");
                //// Devolvemos el archivo como FileResult
                return File(archivoBytes, "application/pdf");


            }
            else
            {
                FilePath = ruta + nombredoc + ".xlsx"; // Ruta al archivo xlsx

                //DESCARGA PARA XLSX

                // Leemos el archivo en bytes
                byte[] archivoBytes = System.IO.File.ReadAllBytes(FilePath);

                // Configuramos el encabezado Content-Disposition para forzar la descarga
                Response.Headers["Content-Disposition"] = "attachment; filename=" + nombredoc + ".xlsx";

                // Configuramos el tipo de contenido
                Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";


                System.IO.File.Delete(FilePath); //se borra el archivo en el servidor

                // Devolvemos el archivo como FileResult
                return File(archivoBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombredoc + ".xlsx");



            }

            
            //HttpResponseMessage response = new HttpResponseMessage();
            //response.StatusCode = HttpStatusCode.OK;


            

            if (System.IO.File.Exists(FilePath))
            {
                // Configuramos el nombre del archivo (puede ser opcional)

                //****EMPIEZA VISUALIZACION*********
                //// Configuramos el encabezado Content-Disposition para visualizar el archivo en el navegador
                //Response.Headers["Content-Disposition"] = "inline; filename=" + nombredoc+".pdf";

                //// Configuramos el tipo de contenido
                //Response.ContentType = "application/pdf";

                //// Leemos el archivo en bytes
                //byte[] archivoBytes = System.IO.File.ReadAllBytes(pdfFilePath);
                ////borro el archivo
                //System.IO.File.Delete(pdfFilePath);
                //// Devolvemos el archivo como FileResult
                //return File(archivoBytes, "application/pdf");

                //****HASTA AQUI VISUALIZACION********

                
            }
            else
            {
                return (HttpStatusCode.NotFound, "El archivo PDF no se encontró.");
            }


            



        }




       













        

    }

    // CODIGO INSERVIBLE QUE A LO MEJOR PUEDE SERVIR EN UN FUTURO NO VAYA A SER

    //FileStream reporte = new FileStream(ruta + nombredoc + ".pdf", FileMode.Open, FileAccess.Read);
    //HttpContext.Response.ContentType = ("application/pdf");
    //string contentDisposition = string.Concat("inline; filename=", new FileInfo(nombredoc + ".pdf").Name);
    //MemoryStream memoria = new MemoryStream();
    //reporte.CopyTo(memoria);
    //reporte.Close();

    //memoria.Position = 0;


    ////Escribimos el Reposte en en content del response 
    //FileResult fileResult = new FileStreamResult(memoria, "application/pdf");
    //fileResult.FileDownloadName = nombredoc + ".pdf";
    //response.Content = new StreamContent(memoria);



    ////Agregar estas dos lineas para descargar directamente el archivo
    //response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse(contentDisposition);

    ////retornar la respuesta
    ///

    //// Crear una instancia de FileResult para la descarga


    //response.Content = new StreamContent(archivoStream);
    //response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
    //response.Content.Headers.ContentDisposition.FileName = "output2.pdf";
    //response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
    //string contentDisposition = string.Concat("attachment; filename=", new FileInfo("output2" + ".pdf").Name);
    //response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse(contentDisposition);

    //return response;


    //// Crear una instancia de FileResult para la descarga


    //response.Content = new StreamContent(archivoStream);
    //response.Content.Headers.ContentDisposition = new ContentDispositionHeaderValue("attachment");
    //response.Content.Headers.ContentDisposition.FileName = "output2.pdf";
    //response.Content.Headers.ContentType = new MediaTypeHeaderValue("application/pdf");
    //string contentDisposition = string.Concat("attachment; filename=", new FileInfo("output2" + ".pdf").Name);
    //response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse(contentDisposition);



    //string contentDisposition = string.Concat("attachment; filename=", new FileInfo("output" + ".pdf").Name);

    //MemoryStream responseStream = new MemoryStream();
    //stream.CopyTo(responseStream);

    // HttpResponseMessage response = new HttpResponseMessage();
    //response.StatusCode = HttpStatusCode.OK;

    // //Escribimos el Reposte en en content del response 
    // response.Content = new StreamContent(responseStream);

    ////Agregar estas dos lineas para descargar directamente el archivo
    //response.Content.Headers.ContentDisposition = ContentDispositionHeaderValue.Parse(contentDisposition);

    // //retornar la respuesta
    // return response;












}

