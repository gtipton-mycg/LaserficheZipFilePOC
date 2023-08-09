//using Microsoft.Extensions.Logging;
using Azure.Core;
//using Services;
using Azure.Identity;
using Azure.Security.KeyVault.Secrets;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Laserfiche_Download_Issues
{
    /// <summary>
    /// This service is registered as a singleton and should not reference other non-singleton services.
    /// </summary>
    public class AzureKeyVaultService : IAzureKeyVaultService
    {
        private readonly AzureKeyVaultConfig _vaultConfig;
        private readonly Dictionary<string, KeyVaultSecret> _keys;
        private SecretClient azureClient;
        private static readonly object _threadLock = new object();

        public AzureKeyVaultService(AzureKeyVaultConfig vaultConfig)
        {
            this._keys = new Dictionary<string, KeyVaultSecret>();
            this._vaultConfig = vaultConfig;
        }

        public async Task<string> GetKey(string key, bool storeForLater = true)
        {
            if (!this._keys.ContainsKey(key))
            {
                KeyVaultSecret secret;

                try
                {
                    // get value from vault
                    var client = this.GetAzureClient();
                    var secretTask = client.GetSecretAsync(key);
                    var secreteResponse = await secretTask;
                    secret = secreteResponse.Value;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error getting value for key: {key} from key vault");
                    Console.WriteLine(ex.Message);
                    throw;
                }

                // store for later use
                if (storeForLater)
                {
                    // ensure that only thread enters this code at a time
                    lock (_threadLock)
                    {
                        // the fetching above is async so let's make sure it's still not in the cache before we add it
                        if (!this._keys.ContainsKey(key))
                        {
                            this._keys.Add(key, secret);
                        }
                    }
                }

                return secret.Value;
            }
            else
                return this._keys[key].Value;
        }

        private SecretClient GetAzureClient()
        {
            if (this.azureClient == null)
            {
                this.azureClient = new SecretClient(new Uri(this._vaultConfig.Vault),
                    new ClientSecretCredential(this._vaultConfig.TenantId, this._vaultConfig.ClientId, this._vaultConfig.ClientSecrete)
                    , new SecretClientOptions()
                    {
                        Retry =
                        {
                            Delay= TimeSpan.FromSeconds(2),
                            MaxDelay = TimeSpan.FromSeconds(16),
                            MaxRetries = 5,
                            Mode = RetryMode.Exponential
                            },
                    });
            }
            return this.azureClient;
        }
    }
}

