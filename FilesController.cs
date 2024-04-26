// ref https://www.c-sharpcorner.com/article/upload-single-and-multiple-files-using-the-net-core-6-web-api/

using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace excel2pdf_webAPI
{
    [Route("excel2pdf")]
    [ApiController]
    public class FilesController : ControllerBase
    {
        [HttpPost("upload")]
        public async Task<IActionResult> OnPostUploadAsync(IFormFile file)
        {
            long size = file.Length;

            //var filePath = Path.GetTempFileName();
            //TODO: check file type, only xlsx file could be upload 
            //TODO: throw error if file had existed in storage

            using (var stream = System.IO.File.Create("./data/" + file.FileName))
            {
                await file.CopyToAsync(stream);
            }

            return Ok(new { filename = file.FileName, size });
        }

        [HttpDelete("{fileName}")]
        public IActionResult DeleteFile(string fileName)
        {
            // Validate and authorize the request (e.g., check user permissions)

            // Construct the full path to the file
            string filePath = Path.Combine(Directory.GetCurrentDirectory(), "data", fileName);

            // Check if the file exists
            if (!System.IO.File.Exists(filePath))
            {
                return NotFound(); // Return 404 if file not found
            }

            //check file type
            var supportedFileTypes = new List<string> { ".pdf", ".xlsx", ".xls" };
            if ( !supportedFileTypes.Contains(Path.GetExtension(filePath)) )
            {
                return StatusCode(415, $"illegal file type of {Path.GetExtension(filePath)}");
            }

            try
            {
                // Attempt to delete the file
                System.IO.File.Delete(filePath);

                return NoContent(); // Return 204 if successful deletion
            }
            catch (Exception ex)
            {
                // Handle deletion error
                return StatusCode(500, $"Error deleting file: {ex.Message}");
            }
        }
    }
}
