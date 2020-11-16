using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.OpenApi.Models;
using OfficeOpenXml;
using Swashbuckle.AspNetCore.SwaggerGen;

namespace ExcelReader.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelReaderController : ControllerBase
    {
        private readonly ILogger<ExcelReaderController> _logger;

        public ExcelReaderController(ILogger<ExcelReaderController> logger)
        {
            _logger = logger;
        }

        [HttpPost("sqlQuery")]
        [FileUploadOperation.FileContentType]

        public IActionResult GetSqlQuery(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File Not Selected");

            string fileExtension = Path.GetExtension(file.FileName);
            if (fileExtension != ".xls" && fileExtension != ".xlsx")
                return BadRequest("File Not Selected");
            
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            var lines = new HashSet<string>();
            var completeSql = "";
            var headerSql =
                @$"INSERT INTO svajoniu_aprasymas ( `vardas`, `aprasymas`, `poreikiai`, `busena`, `photo`, `miestas` ) VALUES ";
            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                var firstSheet = package.Workbook.Worksheets.First();
                int colCount = firstSheet.Dimension.End.Column; //get Column Count
                int rowCount = firstSheet.Dimension.End.Row; //get row count
                var isCorrect = CheckHeaders(firstSheet.Cells);
                if (!isCorrect)
                {
                    return BadRequest("Incorrect headers. Should be: Vardas, Aprasymas, Poreikiai, Busena, Photo, Miestas");
                }
                for (int row = 1; row <= rowCount; row++)
                {
                    for (int col = 1; col <= colCount; col++)
                    {
                        if (row == 1)
                            continue;

                        var vardas = firstSheet.Cells[row, 1].Value?.ToString().Trim();
                        var aprasymas = firstSheet.Cells[row, 2].Value?.ToString().Trim();
                        var poreikiai = firstSheet.Cells[row, 3].Value?.ToString().Trim();
                        var busena = 0;
                        var photo = firstSheet.Cells[row, 5].Value?.ToString().Trim();
                        if (photo == "M")
                        {
                            photo = "pirma.jpg";
                        }
                        else
                        {
                            photo = "antra.jpg";
                        }
                        var miestas = firstSheet.Cells[row, 6].Value?.ToString().Trim();

                        var line = $"('{vardas}', '{aprasymas}', '{poreikiai}', {busena}, '{photo}', '{miestas}' )";
                        lines.Add(line);
                    }
                }

                var insertSql = string.Join(",", lines);
                 completeSql= headerSql + insertSql + ";";

                 return Ok(completeSql);
            }

            bool CheckHeaders(ExcelRange cells)
            {
                var vardas = cells[1, 1].Value?.ToString().Trim();
                var aprasymas = cells[1, 2].Value?.ToString().Trim();
                var poreikiai = cells[1, 3].Value?.ToString().Trim();
                var busena = cells[1, 4].Value?.ToString().Trim();
                var photo = cells[1, 5].Value?.ToString().Trim();
                var miestas = cells[1, 6].Value?.ToString().Trim();

                if (vardas == "Vardas" && aprasymas == "Aprasymas" && poreikiai == "Poreikiai" && busena == "Busena" &&
                    photo == "Photo" && miestas == "Miestas")
                    return true;
                return false;
            }
        }
    }
    
    /// <summary>
    /// Add extra parameters for uploading files in swagger.
    /// </summary>
    public class FileUploadOperation : IOperationFilter
    {
        /// <summary>
        /// Applies the specified operation.
        /// </summary>
        /// <param name="operation">The operation.</param>
        /// <param name="context">The context.</param>
        public void Apply(OpenApiOperation operation, OperationFilterContext context)
        {

            var isFileUploadOperation =
                context.MethodInfo.CustomAttributes.Any(a => a.AttributeType == typeof(FileContentType));

            if (!isFileUploadOperation) return;

            operation.Parameters.Clear();
   
            var uploadFileMediaType = new OpenApiMediaType()
            {
                Schema = new OpenApiSchema()
                {
                    Type = "object",
                    Properties =
                    {
                        ["uploadedFile"] = new OpenApiSchema()
                        {
                            Description = "Upload File",
                            Type = "file",
                            Format = "formData"
                        }
                    },
                    Required = new HashSet<string>(){  "uploadedFile"  }
                }
            };

            operation.RequestBody = new OpenApiRequestBody
            {
                Content = {  ["multipart/form-data"] = uploadFileMediaType   }
            };
        }
    
        /// <summary>
        /// Indicates swashbuckle should consider the parameter as a file upload
        /// </summary>
        [AttributeUsage(AttributeTargets.Method)]
        public class FileContentType : Attribute
        {
       
        }
    }
}