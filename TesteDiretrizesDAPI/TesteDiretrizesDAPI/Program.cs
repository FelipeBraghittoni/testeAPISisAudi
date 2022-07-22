using Microsoft.EntityFrameworkCore;
using TesteDiretrizesDAPI.Models;


var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
builder.Services.AddDbContext<DirDoctosContext>(opt =>
    opt.UseInMemoryDatabase("DirList"));

var app = builder.Build();

// Configure the HTTP request pipeline.
if (builder.Environment.IsDevelopment())
{
    app.UseDeveloperExceptionPage();
 }

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();