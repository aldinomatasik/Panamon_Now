using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;

namespace MonitoringSystem.Controllers
{
    [Route("api/machine")]
    [ApiController]
    public class MachineController : ControllerBase
    {
        private readonly IConfiguration _configuration;

        public MachineController(IConfiguration configuration)
        {
            _configuration = configuration;
        }

        [HttpGet("efficiency")]
        public IActionResult GetMachineEfficiency([FromQuery] string? date)
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
                    SELECT 
                        [MachineName],
                        ISNULL(CAST([OEE] AS VARCHAR), '-') AS [OEE],
                        ISNULL(CAST([OperatingRatio] AS VARCHAR), '-') AS [OperatingRatio],
                        ISNULL(CAST([Ability] AS VARCHAR), '-') AS [Ability],
                        ISNULL(CAST([Quality] AS VARCHAR), '-') AS [Quality]
                    FROM [dbo].[MachineEfficiency]
                    ORDER BY [MachineName]";

                using var cmd = new SqlCommand(query, conn);
                using var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    result.Add(new
                    {
                        machineName = reader["MachineName"]?.ToString() ?? "-",
                        oEE = reader["OEE"]?.ToString() ?? "-",
                        operatingRatio = reader["OperatingRatio"]?.ToString() ?? "-",
                        ability = reader["Ability"]?.ToString() ?? "-",
                        quality = reader["Quality"]?.ToString() ?? "-"
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
    }
}