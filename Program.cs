using Google.Cloud.Firestore;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using OfficeFireSync.Excel;
using OfficeFireSync.Word;
using CommandLine;

namespace OfficeFireSync
{
    class Program
    {
        private static IServiceProvider serviceProvider;

        static async Task Main(string[] args)
        {
            ConfigureServices();

            await CommandLine.Parser.Default.ParseArguments<Options>(args)
              .WithParsedAsync(RunOptions);
        }

        static async Task RunOptions(Options opts)
        {
            if (opts.ExcelType != null)
            {
                if (opts.ExcelType == "default")
                {
                    ProductExcelETL excelETL = serviceProvider.GetService<ProductExcelETL>();
                    await excelETL.SyncToFireStore(opts.InputFile);
                }
                else if (opts.ExcelType == "shopify")
                {
                    ShopifyExcelETL excelETL = serviceProvider.GetService<ShopifyExcelETL>();
                    await excelETL.SyncToFireStore(opts.InputFile);
                }
            }
            else if (opts.WordType != null)
            {
                WordETL wordETL = serviceProvider.GetService<WordETL>();
                wordETL.Extract();
            }
            else
            {
                throw new NotImplementedException("# Unsupported option parsed");
            }
        }

        private static void ConfigureServices()
        {
            var services = new ServiceCollection();
            services.AddTransient<ProductExcelETL, ProductExcelETL>();
            services.AddTransient<ShopifyExcelETL, ShopifyExcelETL>();
            services.AddTransient<WordETL, WordETL>();
            services.AddTransient<ImagePreprocessor, ImagePreprocessor>();

            serviceProvider = services.BuildServiceProvider();
        }

        class Options
        {
            [Option('r', "read", Required = true, HelpText = "Input file to be processed.")]
            public string InputFile { get; set; }

            [Option('e', "excel", Required = false, HelpText = "Type of Excel configuration.")]
            public string ExcelType { get; set; }

            [Option('w', "word", Required = false, HelpText = "Type of Word configuration.")]
            public string WordType { get; set; }
        }
    }
}
