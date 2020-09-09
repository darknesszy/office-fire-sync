using Google.Cloud.Firestore;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeFireSync.Excel;

namespace OfficeFireSync
{
    class Program
    {
        private static IServiceProvider serviceProvider;

        static async Task Main(string[] args)
        {
            ConfigureServices();
            ExcelSyncer excelSyncer = serviceProvider.GetService<ExcelSyncer>();
            await excelSyncer.SyncToFireStore();
        }

        private static void ConfigureServices()
        {
            var services = new ServiceCollection();
            services.AddTransient<ExcelSyncer, ProductExcelSyncer>();
            services.AddTransient<ImagePreprocessor, ImagePreprocessor>();

            serviceProvider = services.BuildServiceProvider();
        }
    }
}
