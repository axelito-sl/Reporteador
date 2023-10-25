using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Excel.Database;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.FileProviders;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddHttpContextAccessor();

//Servicio de conexin a base de datos 
var connectionstring = builder.Configuration.GetConnectionString("princ");
var report = builder.Configuration.GetConnectionString("reporteador");
builder.Services.AddDbContext<ConexionContext>(options =>
{
  //  options.UseMySQL(connectionstring);
    options.UseMySQL(report);

});




var app = builder.Build();





// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseStaticFiles();

app.UseAuthorization();

app.MapControllers();

app.Run();
