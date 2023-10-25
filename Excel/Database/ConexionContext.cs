using DocumentFormat.OpenXml.InkML;
using Excel.Models;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;

namespace Excel.Database
{
    public class ConexionContext:DbContext
    {
        public DbSet<Models.pruebaexcel> pruebaexcel { get; set; }




        //reporteador dinamico
        public DbSet<Models.Planes> planes { get; set; }

        public DbSet<Models.config_tipo_datos> configTipos { get; set; }

        public DbSet<Models.master_reportes> master_reportes { get; set; }

        public DbSet<Models.columnas_reportes> columnas_reportes { get; set; }
        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<columnas_reportes>()
                .HasKey(p => new { p.id_reporte, p.nom_columna });
        }

        //--------------------------------
        public ConexionContext(DbContextOptions<ConexionContext> options) : base(options)
        {
        }
    }
}
