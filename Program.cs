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
            await CommandLine.Parser.Default.ParseArguments<Options>(args).WithParsedAsync(RunOptions);
        }

        static async Task RunOptions(Options opts)
        {
            if (opts.ExcelMode != null)
            {
                switch (opts.ExcelMode)
                {
                    case "table":
                        ProductExcelETL tableBased = serviceProvider.GetService<ProductExcelETL>();
                        await tableBased.SyncToFireStore(opts.InputFile, opts.CollectionName);
                        break;
                    case "sheet":
                        SheetBasedExcelETL sheetBased = serviceProvider.GetService<SheetBasedExcelETL>();
                        await sheetBased.SyncToFireStore(opts.InputFile, opts.CollectionName);
                        break;
                    case "shopify":
                        ShopifyExcelETL shopify = serviceProvider.GetService<ShopifyExcelETL>();
                        await shopify.SyncToFireStore(opts.InputFile, opts.CollectionName);
                        break;
                    default:
                        Console.WriteLine("# Unsupported Excel ETL Model selected!");
                        break;
                }
            }
            else if (opts.WordMode != null)
            {
                if (opts.WordMode == "heading")
                {
                    WordETL wordETL = serviceProvider.GetService<WordETL>();
                    await wordETL.SyncToFirebase(opts.InputFile);
                }
            }
            else
            {
                Console.WriteLine("# Unsupported option parsed!");
            }
        }

        private static void ConfigureServices()
        {
            var services = new ServiceCollection();
            services.AddTransient<ProductExcelETL, ProductExcelETL>();
            services.AddTransient<ShopifyExcelETL, ShopifyExcelETL>();
            services.AddTransient<SheetBasedExcelETL, SheetBasedExcelETL>();
            services.AddTransient<WordETL, WordETL>();
            services.AddTransient<ImagePreprocessor, ImagePreprocessor>();

            serviceProvider = services.BuildServiceProvider();
        }

        class Options
        {
            [Option('r', "read", Required = true, HelpText = "Input file to be processed.")]
            public string InputFile { get; set; }
            [Option('s', "sync", Required = true, HelpText = "Collection name to sync to.")]
            public string CollectionName { get; set; }

            [Option('e', "excel", Required = false, HelpText = "Mode of Excel configuration.")]
            public string ExcelMode { get; set; }

            [Option('w', "word", Required = false, HelpText = "Mode of Word configuration.")]
            public string WordMode { get; set; }
        }
    }
}
