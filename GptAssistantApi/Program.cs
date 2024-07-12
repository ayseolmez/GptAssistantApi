using GptAssistant.Service.AiService;


var builder = WebApplication.CreateBuilder(args);

builder.Services.AddScoped<IAiService, AiService>();

//ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

builder.Services.AddControllers();
builder.Services.AddSingleton<IAiService, AiService>();

builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
