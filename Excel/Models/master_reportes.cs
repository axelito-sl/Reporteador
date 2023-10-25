namespace Excel.Models
{
    public class master_reportes
    {

        public int id { get; set; }
        public string nombre { get; set; }
        public string descripcion { get; set; }
        public int num_columnas { get; set; }

        public string Ordenamiento { get; set; }
        public string Filtros { get; set; }

        public string estatus { get; set; }

    }
}
