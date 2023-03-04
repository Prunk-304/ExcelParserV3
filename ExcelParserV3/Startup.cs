using Microsoft.Extensions.DependencyInjection;
using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.Hosting;


namespace ExcelParserV3;
public class Startup
    {
        public void ConfigureServices(IServiceCollection services)
        {   services.AddSingleton<NLogConfigReader>(new NLogConfigReader("nlogconfig.json"));
            services.AddSingleton<LoggerFactory>();
            services.AddScoped<ColumsListBase>();
            services.AddScoped<ExcelParser>();
            services.AddScoped<PrintToConsole>();
        }

        public void Configure(IApplicationBuilder app, IHostEnvironment env)
        {
        }
    }