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

        [HttpPost("ReadExcelFileWithOutSave")]
        public IActionResult ReadExcelFileWithOutSave([FromForm] IFormFile file)
        {
            try
            {
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest("File is empty");

                // Create a stream from the uploaded file
                using (var stream = file.OpenReadStream())
                {
                    // Create an ExcelDataReader object to read from the stream
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Read each row in the Excel file
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
                                // Example: Read data from Excel and process it
                                var student = new Student();
                                student.Name = reader.GetValue(1).ToString(); // Assuming name is in the second column
                                student.age = Convert.ToInt32(reader.GetValue(2).ToString()); // Assuming age is in the third column

                                // Process the student object as needed, such as saving to a database
                                _ApplicationDbContext.Add(student);
                                _ApplicationDbContext.SaveChanges();
                            }
                        } while (reader.NextResult()); // Move to next result set if any (for multiple sheets in Excel)
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

        [HttpPost("ReadTextWithOutSave")]
        public IActionResult ReadTextWithOutSave([FromForm] IFormFile file)
        {
            try
            {
                //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest();

                using (var reader = new StreamReader(file.OpenReadStream()))
                {
                    var fileContent = reader.ReadToEnd();
                    // Do whatever processing you need with the fileContent here
                    // For example, you can return the file content in the response
                    return Ok(fileContent);
                }
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

        [HttpPost("ReadWordWithOutSave")]
        public IActionResult ReadWordWithOutSave([FromForm] IFormFile file)
        {
            try
            {
                //System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                if (file == null || file.Length == 0)
                    return BadRequest();

                using (var stream = file.OpenReadStream())
                {
                    // Use DocX to read the content of the Word file directly from the stream
                    using (DocX doc = DocX.Load(stream))
                    {
                        string text = doc.Text;
                        // Now you have the text content of the Word file in the 'text' variable
                        return Ok(text);
                    }
                }
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}