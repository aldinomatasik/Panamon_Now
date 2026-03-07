using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace MonitoringSystem.Controllers
{
    [Route("api/machine")]
    [ApiController]
    public class MachineController : ControllerBase
    {
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public MachineController(IConfiguration configuration, IWebHostEnvironment webHostEnvironment)
        {
            _configuration = configuration;
            _webHostEnvironment = webHostEnvironment;
        }

        [HttpGet("efficiency")]
        public IActionResult GetMachineEfficiency(
            [FromQuery] int month,
            [FromQuery] int year)
        {
            if (month < 1 || month > 12)
                return BadRequest(new { error = "Bulan tidak valid (1–12)." });

            if (year < 2000 || year > 2100)
                return BadRequest(new { error = "Tahun tidak valid." });

            var result = new List<object>();
            try
            {
                var connStr = _configuration.GetConnectionString("DefaultConnection");
                if (string.IsNullOrEmpty(connStr))
                    return StatusCode(500, new { error = "Connection string tidak ditemukan." });

                using var conn = new SqlConnection(connStr);
                conn.Open();

                var query = @"
                    SELECT 
                        [MachineName],
                        ISNULL(CAST(ROUND(AVG(CAST([OEE]            AS FLOAT)), 2) AS VARCHAR), '-') AS [OEE],
                        ISNULL(CAST(ROUND(AVG(CAST([OperatingRatio] AS FLOAT)), 2) AS VARCHAR), '-') AS [OperatingRatio],
                        ISNULL(CAST(ROUND(AVG(CAST([Ability]        AS FLOAT)), 2) AS VARCHAR), '-') AS [Ability],
                        ISNULL(CAST(ROUND(AVG(CAST([Quality]        AS FLOAT)), 2) AS VARCHAR), '-') AS [Quality],
                        ISNULL(CAST(ROUND(AVG(CAST([Achievement]    AS FLOAT)), 2) AS VARCHAR), '-') AS [Achievement]
                    FROM [dbo].[MachineEfficiency]
                    WHERE MONTH([Date]) = @Month
                      AND YEAR([Date])  = @Year
                    GROUP BY [MachineName]
                    ORDER BY [MachineName]";

                using var cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@Month", month);
                cmd.Parameters.AddWithValue("@Year", year);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        machineName = reader["MachineName"]?.ToString() ?? "-",
                        oEE = reader["OEE"]?.ToString() ?? "-",
                        operatingRatio = reader["OperatingRatio"]?.ToString() ?? "-",
                        ability = reader["Ability"]?.ToString() ?? "-",
                        quality = reader["Quality"]?.ToString() ?? "-",
                        achievement = reader["Achievement"]?.ToString() ?? "-"
                    });
                }

                return Ok(result);
            }
            catch (SqlException sqlEx)
            {
                return StatusCode(500, new { error = $"Database error: {sqlEx.Message}" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpGet("list")]
        public IActionResult GetMachineList()
        {
            var result = new List<object>();
            try
            {
                var connStr = _configuration.GetConnectionString("DefaultConnection");
                if (string.IsNullOrEmpty(connStr))
                    return StatusCode(500, new { error = "Connection string tidak ditemukan." });

                using var conn = new SqlConnection(connStr);
                conn.Open();

                var query = @"
                    SELECT DISTINCT [MachineName]
                    FROM [dbo].[MachineEfficiency]
                    WHERE [MachineName] IS NOT NULL AND [MachineName] != ''
                    ORDER BY [MachineName]";

                using var cmd = new SqlCommand(query, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    result.Add(new
                    {
                        machineName = reader["MachineName"]?.ToString() ?? "-"
                    });
                }

                return Ok(result);
            }
            catch (SqlException sqlEx)
            {
                return StatusCode(500, new { error = $"Database error: {sqlEx.Message}" });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = ex.Message });
            }
        }

        [HttpGet("download-template")]
        public IActionResult DownloadTemplate()
        {
            var filePath = Path.Combine(
                _webHostEnvironment.WebRootPath,
                "data", "MachineEfficiency", "Machine_input_template.xlsx"
            );

            if (!System.IO.File.Exists(filePath))
                return NotFound(new { error = "File template tidak ditemukan di server." });

            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "Machine_input_template.xlsx");
        }

        [HttpPost("import")]
        public async Task<IActionResult> ImportMachineEfficiency(
            [FromForm] IFormFile file,
            [FromForm] string machineName,
            [FromForm] int month,
            [FromForm] int year)
        {
            if (file == null || file.Length == 0)
                return BadRequest(new { error = "File tidak boleh kosong." });

            if (string.IsNullOrEmpty(machineName))
                return BadRequest(new { error = "Nama machine harus dipilih." });

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var newData = new List<(double? Achievement, double? OperatingRatio, double? Quality, double? Ability, double? OEE, DateTime Date)>();

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);

                using var package = new ExcelPackage(stream);
                var sheet = package.Workbook.Worksheets[0];
                int rowCount = sheet.Dimension.Rows;

                for (int row = 3; row <= rowCount; row++)
                {
                    if (sheet.Cells[row, 3].Value == null) continue;

                    for (int day = 1; day <= 31; day++)
                    {
                        if (day > DateTime.DaysInMonth(year, month)) break;

                        int baseCol = 3 + (day - 1) * 4;

                        double? achievement = TryParseDouble(sheet.Cells[row, baseCol].Value);
                        double? operatingRatio = TryParseDouble(sheet.Cells[row, baseCol + 1].Value);
                        double? quality = TryParseDouble(sheet.Cells[row, baseCol + 2].Value);
                        double? ability = TryParseDouble(sheet.Cells[row, baseCol + 3].Value);

                        if ((achievement == null || achievement == 0) &&
                            (operatingRatio == null || operatingRatio == 0) &&
                            (quality == null || quality == 0) &&
                            (ability == null || ability == 0))
                            continue;

                        double? oee = null;
                        if (operatingRatio.HasValue && ability.HasValue && quality.HasValue)
                        {
                            oee = Math.Round(
                                (operatingRatio.Value / 100.0) *
                                (ability.Value / 100.0) *
                                (quality.Value / 100.0) * 100.0, 2);
                        }

                        newData.Add((
                            Achievement: achievement,
                            OperatingRatio: operatingRatio,
                            Quality: quality,
                            Ability: ability,
                            OEE: oee,
                            Date: new DateTime(year, month, day)
                        ));
                    }
                }

                var connStr = _configuration.GetConnectionString("DefaultConnection");
                using var conn = new SqlConnection(connStr);
                conn.Open();

                using (var deleteCmd = new SqlCommand(@"
                    DELETE FROM [dbo].[MachineEfficiency]
                    WHERE [MachineName] = @MachineName
                      AND MONTH([Date]) = @Month
                      AND YEAR([Date])  = @Year", conn))
                {
                    deleteCmd.Parameters.AddWithValue("@MachineName", machineName);
                    deleteCmd.Parameters.AddWithValue("@Month", month);
                    deleteCmd.Parameters.AddWithValue("@Year", year);
                    deleteCmd.ExecuteNonQuery();
                }

                int insertedCount = 0;
                foreach (var item in newData)
                {
                    using var insertCmd = new SqlCommand(@"
                        INSERT INTO [dbo].[MachineEfficiency]
                            ([MachineName], [Achievement], [OperatingRatio], [Quality], [Ability], [OEE], [Date])
                        VALUES
                            (@MachineName, @Achievement, @OperatingRatio, @Quality, @Ability, @OEE, @Date)", conn);

                    insertCmd.Parameters.AddWithValue("@MachineName", machineName);
                    insertCmd.Parameters.AddWithValue("@Achievement", (object?)item.Achievement ?? DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@OperatingRatio", (object?)item.OperatingRatio ?? DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@Quality", (object?)item.Quality ?? DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@Ability", (object?)item.Ability ?? DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@OEE", (object?)item.OEE ?? DBNull.Value);
                    insertCmd.Parameters.AddWithValue("@Date", item.Date);

                    insertCmd.ExecuteNonQuery();
                    insertedCount++;
                }

                return Ok(new
                {
                    success = true,
                    message = $"Berhasil import {insertedCount} data untuk {machineName} (Bulan {month}/{year})."
                });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { error = "Gagal Import: " + ex.Message });
            }
        }

        private double? TryParseDouble(object? val)
        {
            if (val == null) return null;
            if (double.TryParse(val.ToString(), out double result))
                return result;
            return null;
        }
    }
}