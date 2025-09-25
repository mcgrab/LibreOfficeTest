using LibreOfficeTest.HealthChecks;
using LibreOfficeTest.HostedServices;
using Microsoft.AspNetCore.Diagnostics.HealthChecks;

namespace LibreOfficeTest
{
    public class Program
    {
        public static void Main(string[] args)
        {
            var builder = WebApplication.CreateBuilder(args);

            // Add services to the container.

            builder.Services.AddControllers();
            // Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
            builder.Services.AddEndpointsApiExplorer();
            builder.Services.AddSwaggerGen();
            builder.Services.AddHostedService<LibreOfficeHostedService>();

            builder.Services.AddHealthChecks()
                .AddCheck<LibreOfficeUnoHealthCheck>(
                name:"libreoffice-uno",
                tags: HealthChecksTags.ReadinessTags);

            var app = builder.Build();

            app.UseSwagger(options =>
            {
                options.RouteTemplate = "api/{documentName}/swagger.json";
            });
            app.UseSwaggerUI(options =>
            {
                options.RoutePrefix = "api";
                options.SwaggerEndpoint("/api/v1/swagger.json", "Converter");
            });

            app.UseRouting();

            app.UseHttpsRedirection();

            app.UseAuthorization();

            app.MapControllers();

            app.UseEndpoints(endpoints =>
            {
                endpoints.MapHealthChecks("/health/live", new HealthCheckOptions { Predicate = check => check.Tags.Contains(HealthChecksTags.LivenessTag) });
                endpoints.MapHealthChecks("/health/ready", new HealthCheckOptions { Predicate = check => check.Tags.Contains(HealthChecksTags.ReadinessTag) });
            });

            app.Run();
        }
    }
}
