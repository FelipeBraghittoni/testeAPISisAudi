using Microsoft.EntityFrameworkCore;
using TesteDiretrizesDAPI.Models;

var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddCors(options =>
{
    options.AddPolicy(name: MyAllowSpecificOrigins,
                      builder =>
                      {
                          builder.WithOrigins("*").AllowAnyMethod().AllowAnyHeader();
                          //alterar * pelas urls de acesso
                      });
});


// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddDbContext<DirDoctosContext>(opt =>
    opt.UseInMemoryDatabase("DirList"));
builder.Services.AddDbContext<LoginUsuariosContext>(opt =>
    opt.UseInMemoryDatabase("UserList"));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (builder.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
 }

app.UseCors(MyAllowSpecificOrigins);


app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();