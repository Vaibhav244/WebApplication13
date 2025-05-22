using Microsoft.EntityFrameworkCore;
using WebApplication13.Context;
using WebApplication13.Repositories.Implementations;
using WebApplication13.Repositories.Interfaces;
using WebApplication13.Services.Implementations;
using WebApplication13.Services.Interfaces;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllers();
builder.Services.AddDbContext<MyDbContext>(options =>
    options.UseOracle(builder.Configuration.GetConnectionString("DefaultConnection")));
builder.Services.AddControllers().AddJsonOptions(options => options.JsonSerializerOptions.PropertyNamingPolicy = null);

// Register repositories
builder.Services.AddScoped<IEmployeeRepository, EmployeeRepository>();
builder.Services.AddScoped<IAttendanceRepository, AttendanceRepository>();
builder.Services.AddScoped<IMailLogRepository, MailLogRepository>();
builder.Services.AddScoped<IOnsiteEmployeeRepository, OnsiteEmployeeRepository>();

// Register services
builder.Services.AddScoped<IWFOReportService, WFOReportService>();
builder.Services.AddScoped<IWFHReportService, WFHReportService>();
builder.Services.AddScoped<IMailLogService, MailLogService>();
builder.Services.AddScoped<IOnsiteEmployeeService, OnsiteEmployeeService>();

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();