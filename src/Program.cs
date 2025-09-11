using DataMerger.Interfaces;
using DataMerger.Services;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

class Program
{
    static async Task Main(string[] args)
    {
        using var host = Host.CreateDefaultBuilder(args)
            .ConfigureServices((context, services) =>
            {
                // Registrar serviços
                services.AddScoped<IDataMergeService, DataMergeService>();
            })
            .Build();

        // Resolver serviço e executar
        var service = host.Services.GetRequiredService<IDataMergeService>();
        await service.StartProcess();
    }
}
