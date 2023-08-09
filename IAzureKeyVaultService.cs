using System.Threading.Tasks;

namespace Laserfiche_Download_Issues
{
    public interface IAzureKeyVaultService
    {
        Task<string> GetKey(string key, bool storeForLater = true);
    }
}
