using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using OfficeOpenXml;

namespace ExcelReader.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EldersReaderController : ControllerBase
    {
        private readonly ILogger<EldersReaderController> _logger;

        public EldersReaderController(ILogger<EldersReaderController> logger)
        {
            _logger = logger;
        }

        [HttpPost("eldersFile")]
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
            var headerSql = @"INSERT INTO `Senoliai` (`vardas`, `pavarde`, `amzius`, `lytis`, `adresas`, `telefonas`, `soc_darbuotojas`, `2020_poreikis`, `2021_poreikis`) VALUES ";
            using (var package = new ExcelPackage(file.OpenReadStream()))
            {
                var firstSheet = package.Workbook.Worksheets.First();
                int colCount = firstSheet.Dimension.End.Column; //get Column Count
                int rowCount = firstSheet.Dimension.End.Row;

                for (int row = 1; row <= rowCount; row++)
                {
                    // skip eil.nr ir atsakinga asmeni
                    for (int col = 3; col <= colCount; col++)
                    {
                        if (row == 1)
                            continue;

                        var vardas = firstSheet.Cells[row, 3].Value?.ToString().Trim();
                        var pavarde = firstSheet.Cells[row, 4].Value?.ToString().Trim();
                        var ageInput = firstSheet.Cells[row, 5].Value?.ToString().Trim();
                        var lytis = firstSheet.Cells[row, 6].Value?.ToString().Trim();
                        var adresas = firstSheet.Cells[row, 7].Value?.ToString().Trim();
                        var telefonas = firstSheet.Cells[row, 8].Value?.ToString().Trim();
                        var darbuotojas = firstSheet.Cells[row, 9].Value?.ToString().Trim();
                        var poreikisOld = firstSheet.Cells[row, 10].Value?.ToString().Trim();
                        var poreikisNew = firstSheet.Cells[row, 11].Value?.ToString().Trim();
                        var importuota = firstSheet.Cells[row, 12].Value?.ToString().Trim();

                        ageInput ??= "0";
                        var senolioAmzius = Convert.ToInt32(Regex.Replace(
                            ageInput, "[^0-9]", // Select everything that is not in the range of 0-9
                                    ""        // Replace that with an empty string.
                                ));

                        var line = $"('{vardas}', '{pavarde}', {senolioAmzius}, '{lytis}', '{adresas}', '{telefonas}', '{darbuotojas}', '{poreikisOld}', '{poreikisNew}')";
                        if (vardas != null && pavarde != null && importuota.ToLower() == "ne")
                            lines.Add(line);
                    }
                }
            }

            var insertSql = string.Join(",", lines);
            completeSql= headerSql + insertSql + ";";

            return Ok(completeSql);
        }
    }
}