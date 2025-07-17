using Microsoft.EntityFrameworkCore;
using BestPolicyReport.Data;
using BestPolicyReport.Services.DailyPolicyService;
using BestPolicyReport.Services.BillService;
using BestPolicyReport.Services.CashierService;
using BestPolicyReport.Services.OutputVatService;
using BestPolicyReport.Services.ArApService;
using BestPolicyReport.Services.PremInDirectService;
using BestPolicyReport.Services.WhtService;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using Microsoft.AspNetCore.DataProtection;
using System.IdentityModel.Tokens.Jwt;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Serilog;
using Serilog.Events;




Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Debug()
    .WriteTo.Console()
    .WriteTo.File(
    path: "Logs/log-.txt",
    //outputTemplate: "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] {MethodName} {Message:lj}{NewLine}{Exception}",
    rollOnFileSizeLimit : true,
    fileSizeLimitBytes :  1 * 1024 * 1024,
    rollingInterval: RollingInterval.Day,
    restrictedToMinimumLevel: LogEventLevel.Information
    ).Enrich.FromLogContext()
                .Enrich.WithProperty("ClassName", typeof(Program).Name) // Add Class Name
                .Enrich.WithProperty("MethodName", () =>
                {
                    var stackFrame = new System.Diagnostics.StackFrame(1);
                    return stackFrame.GetMethod().Name;
                })
               
    .CreateLogger();


try
{
    Log.Information("Starting web application");
    var builder = WebApplication.CreateBuilder(args);

    //Jwt configuration starts here
    var jwtKey = builder.Configuration["Jwt:Key"].ToString();
    var key = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey));
    var creds = new SigningCredentials(key, SecurityAlgorithms.HmacSha256Signature);

    builder.Services.AddAuthentication(cfg =>
{
    cfg.DefaultAuthenticateScheme = JwtBearerDefaults.AuthenticationScheme;
    cfg.DefaultChallengeScheme = JwtBearerDefaults.AuthenticationScheme;
    cfg.DefaultScheme = JwtBearerDefaults.AuthenticationScheme;
})
 .AddJwtBearer(options =>
 {
     options.TokenValidationParameters = new TokenValidationParameters
     {
         ValidateIssuer = false,
         ValidateAudience = false,
         ValidateLifetime = true,
         ValidateIssuerSigningKey = true,
         //ValidIssuer = jwtIssuer,
         //ValidAudience = jwtIssuer,
         //IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(jwtKey))
         IssuerSigningKey = key
     };

     options.Events = new JwtBearerEvents
     {
         OnTokenValidated = context =>
         {
             var jwtToken = context.SecurityToken as JwtSecurityToken;
             if (jwtToken == null)
             {
                 context.Fail("Invalid token");
                 return Task.CompletedTask;
             }

             var usernameClaim = jwtToken.Claims.FirstOrDefault(c => c.Type == "USERNAME")?.Value;
             var roleClaim = jwtToken.Claims.FirstOrDefault(c => c.Type == "ROLE")?.Value;

             if (string.IsNullOrEmpty(usernameClaim) || string.IsNullOrEmpty(roleClaim))
             {
                 context.Fail("Invalid token claims");
                 return Task.CompletedTask;
             }

             // Additional custom validation logic for the claims
             if (roleClaim != "admin" && roleClaim != "account")
             {
                 context.Fail("Invalid role");
                 return Task.CompletedTask;
             }

             return Task.CompletedTask;
         }
     };
 });


//Jwt configuration ends here


builder.Services.AddCors(options =>
{
    options.AddDefaultPolicy(builder =>
    {
        builder.AllowAnyOrigin()
               .AllowAnyMethod()
               .AllowAnyHeader();
    });
});



// Add services to the container.

//builder.Services.AddControllers(); // old
builder.Services.AddControllersWithViews();// new

// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddScoped<IOICService, OICService>();
builder.Services.AddScoped<IBillService, BillService>();
builder.Services.AddScoped<ICashierService, CashierService>();
builder.Services.AddScoped<IOutputVatService, OutputVatService>();
builder.Services.AddScoped<IArApService, ArApService>();
builder.Services.AddScoped<IPremInDirectService, PremInDirectService>();
builder.Services.AddScoped<IWhtService, WhtService>();

builder.Services.AddScoped<IOICService,OICService>();
    builder.Services.AddDbContext<DataContext>(options => options.UseNpgsql(builder.Configuration.GetConnectionString("Amitydb")));
//builder.Services.AddLogging();
builder.Services.AddSerilog(); // <-- Add this line
    var app = builder.Build();

// log
//var logFilePath = Path.Combine(AppContext.BaseDirectory, builder.Configuration["Logging:LogFilePath"].ToString());
//var logFilePath = "\\app\\Logs\\log-{Date}.txt";

//var logFilePath = builder.Configuration["Logging:LogFilePath"].ToString();
    
//var loggerFactory = app.Services.GetService<ILoggerFactory>();
//loggerFactory.AddFile(logFilePath);


// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

    app.UseSerilogRequestLogging(); // <-- Add this line
    app.UseCors();

app.UseHttpsRedirection();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllers();

app.Run();

}
catch (Exception ex)
{
    Log.Fatal(ex, "Application terminated unexpectedly");
}
finally
{
    Log.CloseAndFlush();
}

