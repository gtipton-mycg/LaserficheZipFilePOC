// See https://aka.ms/new-console-template for more information

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.IO;

namespace Laserfiche_Download_Issues
{
    public class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("*** User Notes ***");
            Console.WriteLine("Demonstrate issues with the download functionality in the Laserfiche SDK.");
            Console.WriteLine("The issue we have found is when downloading TIFF PDF files that contain digital annotation stamps.");
            Console.WriteLine("These digital stamps are public images that are used to overlay text and signatures on the PDF files.");
            Console.WriteLine("These stamps are applied by Laserfiche repository users.");
            Console.WriteLine("*** End Of User Notes ***");
            Console.WriteLine("--");

            Console.WriteLine("Constructing dependency injection via appsettings.json");

            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            IConfigurationRoot configuration = builder.Build();

            // Create service collection
            ServiceCollection services = new ServiceCollection();

            // configure strongly typed settings objects
            //var appSettingsSection = configuration.GetSection("AppSettings");
            //services.Configure<AppSettings>(app_settings => { appSettingsSection.Bind(app_settings); });
            //var appSettings = appSettingsSection.Get<AppSettings>();

            IConfigurationSection azureKeyVaultConfigSection = configuration.GetSection("AzureKeyVault");
            services.Configure<AzureKeyVaultConfig>(config => { azureKeyVaultConfigSection.Bind(config); });
            AzureKeyVaultConfig azureKeyVaultConfig = azureKeyVaultConfigSection.Get<AzureKeyVaultConfig>();
            Console.WriteLine("AzureKeyVaultConfig.Vault: " + azureKeyVaultConfig.Vault);

            IConfigurationSection laserficheConfigSection = configuration.GetSection("Laserfiche");
            services.Configure<LaserficheConfig>(config => { laserficheConfigSection.Bind(config); });
            LaserficheConfig laserficheConfig = laserficheConfigSection.Get<LaserficheConfig>();
            Console.WriteLine("laserficheConfig.LaserficheRepoName: " + laserficheConfig.LaserficheRepoName);

            ConfigureServices(services, configuration, azureKeyVaultConfig, laserficheConfig);

            // Create service provider
            IServiceProvider serviceProvider = services.BuildServiceProvider();

            Console.WriteLine("Finished constructing dependency injection via appsettings.json");

            Console.WriteLine("Calling the Laserfiche Download sample code");
            CallDocumentDownload(serviceProvider);
            Console.WriteLine("Finished calling the Laserfiche Download sample code");
            Console.WriteLine("Please close this window or click ctrl + c");

            // Pause the Console so we can read the output.
            // Click ctrl + c or close the open dialog to exit the program
            Console.ReadLine(); // pause after completion
        }

        private static void ConfigureServices(IServiceCollection services, IConfigurationRoot configuration, AzureKeyVaultConfig azureKeyVaultConfig, LaserficheConfig laserficheConfig)
        {
            services.AddMemoryCache();

            ////////////////////////////////////////////////////
            //// AzureKeyVault Config
            ////////////////////////////////////////////////////
            //var vaultConfig = new AzureKeyVaultConfig();
            //configuration.Bind("AzureKeyVault", vaultConfig);
            services.AddSingleton(azureKeyVaultConfig);
            ////////////////////////////////////////////////////

            ////////////////////////////////////////////////////
            //// LF Config
            ////////////////////////////////////////////////////
            //var lfConfig = new LaserficheConfig();
            //configuration.Bind("Laserfiche", lfConfig);
            services.AddSingleton(laserficheConfig);
            ////////////////////////////////////////////////////

            // !!! DO NOT PUT DI CODE ABOVE HERE !!! //
            // Key Vault
            services.AddSingleton<IAzureKeyVaultService, AzureKeyVaultService>();
            // Add app
            services.AddScoped<IDocumentDownload, DocumentDownload>();
            // !!! DO NOT PUT DI CODE BELOW HERE !!! //
        }

        static void CallDocumentDownload(IServiceProvider serviceProvider)
        {
            //using IServiceScope serviceScope = serviceProvider.CreateScope();
            //IServiceProvider provider = serviceScope.ServiceProvider;
            AzureKeyVaultConfig keyVaultConfig = serviceProvider.GetRequiredService<AzureKeyVaultConfig>();
            IAzureKeyVaultService keyVaultService = serviceProvider.GetRequiredService<IAzureKeyVaultService>();

            LaserficheConfig laserficheConfig = serviceProvider.GetRequiredService<LaserficheConfig>();

            IDocumentDownload injectedDownloader = serviceProvider.GetRequiredService<IDocumentDownload>();
            //IDocumentDownload nonInjectedDocumentDownload = new DocumentDownload(keyVaultService, laserficheConfig);
            injectedDownloader.DownloadLaserficheDocument();
            Console.WriteLine();
        }
    }
}