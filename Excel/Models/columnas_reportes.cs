using Microsoft.EntityFrameworkCore;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace Excel.Models
{
    
    public class columnas_reportes
    {
        [Key]
        public int id_reporte { get; set; }
        [Key]
        public string nom_columna { get; set; }
        public string nombre_tipo_dato { get; set; }
        public string operaciones_disp { get; set; }
        public string estatus { get; set; }
    }

}
