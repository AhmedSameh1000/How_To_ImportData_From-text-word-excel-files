using Data.Data;
using Data.Models;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using UdemyProject.Application.ServicesImplementation.FileServiceImplementation;
using Xceed.Words.NET;

namespace Data.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class StudentController : ControllerBase
    {
        private readonly ApplicationDbContext _ApplicationDbContext;
        private readonly IWebHostEnvironment _WebHostEnvironment;
        private readonly IFileService _FileService;

        public StudentController(
            ApplicationDbContext applicationDbContext,
            IWebHostEnvironment webHostEnvironment,
            IFileService fileService)
        {
            _ApplicationDbContext = applicationDbContext;
            _WebHostEnvironment = webHostEnvironment;
            _FileService = fileService;
        }

        [HttpPost("ReadExcelFile")]
        public IActionResult ReadExcel([FromForm] IFormFile file)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest();

                var FolderUploads = Path.Combine(_WebHostEnvironment.WebRootPath, "uploads");

                var fileResult = _FileService.SaveFile(file, FolderUploads);

                using (var Reader = System.IO.File.Open(Path.Combine(FolderUploads, fileResult.Path), FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    using (var reader = ExcelReaderFactory.CreateReader(Reader))
                    {
                        // Choose one of either 1 or 2:

                        var isHeaderSkipped = false;
                        do
                        {
                            while (reader.Read())
                            {
                                if (!isHeaderSkipped)
                                {
                                    isHeaderSkipped = true;
                                    continue;
                                }

                                var Student = new Student();
                                Student.Name = reader.GetValue(1).ToString();
                                Student.age = Convert.ToInt32(reader.GetValue(2).ToString());

                                _ApplicationDbContext.Add(Student);
                                _ApplicationDbContext.SaveChanges();
                            }
                        } while (reader.NextResult());
                    }
                }
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("ReadText")]
        public IActionResult ReadText([FromForm] IFormFile file)
        {
            try
            {
                //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest();

                var FolderUploads = Path.Combine(_WebHostEnvironment.WebRootPath, "uploads");

                var filePath = _FileService.SaveFile(file, FolderUploads).Path;

                if (!Path.Exists(Path.Combine(FolderUploads, filePath)))
                {
                    return BadRequest();
                }
                var Reader = System.IO.File.ReadAllText(Path.Combine(FolderUploads, filePath));

                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpPost("ReadWord")]
        public IActionResult ReadWord([FromForm] IFormFile file)
        {
            try
            {
                //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest();

                var FolderUploads = Path.Combine(_WebHostEnvironment.WebRootPath, "uploads");

                var FilePath = _FileService.SaveFile(file, FolderUploads).Path;

                if (!Path.Exists(Path.Combine(FolderUploads, FilePath)))
                {
                    return BadRequest();
                }

                // Use DocX to read the content of the Word file
                using (DocX doc = DocX.Load(Path.Combine(FolderUploads, FilePath)))
                {
                    string text = doc.Text;
                    // Now you have the text content of the Word file in the 'text' variable
                    return Ok(text);
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}