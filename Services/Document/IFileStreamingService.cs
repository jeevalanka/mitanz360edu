using Microsoft.AspNetCore.Mvc;

namespace MITANZ360Edu.Web.Services.DocumentProcessing
{
    public interface IFileStreamingService
    {
        Task<FileStreamResult> GetFileAsync(
            string itemId,
            CancellationToken cancellationToken = default);
    }
}
