using Data.Services;

namespace UdemyProject.Application.ServicesImplementation.FileServiceImplementation
{
    public interface IFileService
    {
        FileInformation SaveFile(IFormFile file, string FolderPath);
    }
}