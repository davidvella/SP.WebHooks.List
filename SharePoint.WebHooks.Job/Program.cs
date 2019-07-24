using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Hosting;

namespace SharePoint.WebHooks.Job
{
    // To learn more about Microsoft Azure WebJobs SDK, please see https://go.microsoft.com/fwlink/?LinkID=320976
    class Program
    {
        // Please set the following connection strings in app.config for this WebJob to run:
        // AzureWebJobsDashboard and AzureWebJobsStorage
        public static async Task Main(string[] args)
        {
            var builder = new HostBuilder()
                            .ConfigureAppConfiguration((hostingContext, config) =>
                            {
                                config.SetBasePath(Directory.GetCurrentDirectory())
                                .AddJsonFile("appSettings.json", optional: false, reloadOnChange: true)
                                .AddEnvironmentVariables();
                            }
                            )
                            .UseEnvironment("Development")
                            .ConfigureWebJobs(b =>
                            {
                                b.AddAzureStorage();
                            })
                            .UseConsoleLifetime();

            var host = builder.Build();
            using (host)
            {
                await host.RunAsync();
            }
        }
    }
}
