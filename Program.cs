using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using SharepointDailyDigest.Services;

try
{
    var host = new HostBuilder()
        .ConfigureFunctionsWorkerDefaults()
        .ConfigureServices(services =>
        {
            var appInsightsConnectionString = Environment.GetEnvironmentVariable("APPLICATIONINSIGHTS_CONNECTION_STRING");
            if (!string.IsNullOrEmpty(appInsightsConnectionString))
            {
                try
                {
                    services.AddApplicationInsightsTelemetryWorkerService();
                    services.ConfigureFunctionsApplicationInsights();
                }
                catch (Exception aiEx)
                {
                    // Don't let bad App Insights config kill the worker; log and continue
                    Console.Error.WriteLine($"[Application Insights] Skipped: {aiEx.Message}");
                }
            }
            services.AddSingleton<ISharePointDigestService>(_ => new SharePointDigestService());
            services.AddSingleton<IEmailService>(_ => new EmailService());
        })
        .Build();

    await host.RunAsync();
}
catch (Exception ex)
{
    // Log to stderr so Azure Log stream / Kudu shows the real cause of WorkerProcessExitException
    await Console.Error.WriteLineAsync($"[Worker startup failed] {ex}");
    if (ex is AggregateException agg)
        foreach (var inner in agg.InnerExceptions)
            await Console.Error.WriteLineAsync($"[Inner] {inner}");
    if (ex.InnerException != null)
        await Console.Error.WriteLineAsync($"[Inner] {ex.InnerException}");
    throw;
}
