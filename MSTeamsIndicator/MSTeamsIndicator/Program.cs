using MSTeamsIndicator;

using Microsoft.Extensions.Hosting;
using Serilog;

Log.Logger = new LoggerConfiguration()
    .MinimumLevel.Information()
    .WriteTo.File(
        path: "logs\\MSTeamsIndicator-.log",
        rollingInterval: RollingInterval.Day,
        retainedFileCountLimit: 14,
        shared: true)
    .CreateLogger();
try
{
   Log.Information("Starting TeamsSerialWorker");

   Host.CreateDefaultBuilder(args)
      .UseWindowsService(opt =>
      {
         opt.ServiceName = "MSTeamsIndicator";
      })
      .UseSerilog()
      .ConfigureServices(services =>
      {
         services.AddHostedService<Worker>();
      })
      .Build()
      .Run();
}
catch (Exception ex)
{
   Log.Fatal(ex, "Service terminated unexpectedly");
}
finally
{
   Log.CloseAndFlush();
}