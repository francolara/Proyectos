using System;
using System.IO;
using System.Security.Cryptography;
using System.Threading;
using System.Threading.Tasks;

namespace SistemaVisual.Services
{
    public sealed class HashService
    {
        public Task<string> CalcularSha256Async(string ruta, CancellationToken cancellationToken)
        {
            return Task.Run(() =>
            {
                cancellationToken.ThrowIfCancellationRequested();
                using (var stream = new FileStream(ruta, FileMode.Open, FileAccess.Read, FileShare.Read, 81920, FileOptions.SequentialScan))
                using (var sha256 = SHA256.Create())
                {
                    var hash = sha256.ComputeHash(stream);
                    cancellationToken.ThrowIfCancellationRequested();
                    return BitConverter.ToString(hash).Replace("-", string.Empty).ToLowerInvariant();
                }
            }, cancellationToken);
        }
    }
}
