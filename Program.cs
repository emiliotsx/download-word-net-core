var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

// Definir política de CORS
builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyHeader()
              .AllowAnyMethod();
    });
});

var app = builder.Build();

// Habilitar la política de CORS
app.UseCors("AllowAll");

app.UseRouting();
app.MapControllers();
app.Run();
