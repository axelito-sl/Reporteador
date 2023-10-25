using System.ComponentModel.DataAnnotations.Schema;
using System.ComponentModel.DataAnnotations;

namespace Excel.Models
{
    public class pruebaexcel
    {
        [Key]
        [DatabaseGenerated(DatabaseGeneratedOption.Identity)]
        public int id_cliente { get; set; }


        public string nombre_cliente { get; set; }

       public decimal saldo_inicial { get; set; }

        public decimal enero { get; set; }
        public decimal febrero { get; set; }
        public decimal marzo { get; set; }
        public decimal abril { get; set; }
        public decimal mayo { get; set; }
        public decimal junio { get; set; }

       
    }
}
