using ApiConsumoOutlook.Data;
using Microsoft.EntityFrameworkCore;

namespace ApiConsumoOutlook
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.
            builder.Services.AddAuthorization();

            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();

            var conexion = builder.Configuration.GetConnectionString("MiConexion");

            builder.Services.AddDbContext<ApplicationDbContext>(options => options.UseSqlServer(conexion));

            var app = builder.Build();

            // Configure the HTTP request pipeline.
            if (app.Environment.IsDevelopment())
            {
                app.UseSwagger();
                app.UseSwaggerUI();
            }

            app.UseHttpsRedirection();

            app.UseAuthorization();

            

            app.MapGet("/proyectos", async (HttpContext httpContext, ApplicationDbContext context) =>
            {
                return await context.Proyectos.ToListAsync();
            })
            .WithName("Proyecto");

            app.Run();
        }
    }
}