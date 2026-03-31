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

        // ─── Helper: buka koneksi ke MachineDB ───────────────────────────────────
        private SqlConnection OpenConn()
        {
            var connStr = _configuration.GetConnectionString("MachineConnection");
            if (string.IsNullOrEmpty(connStr))
                throw new InvalidOperationException("Connection string 'MachineConnection' tidak ditemukan.");
            var conn = new SqlConnection(connStr);
            conn.Open();
            return conn;
        }

        // ─── GET /api/machine/efficiency?month=X&year=Y ───────────────────────────
        // Ambil rata-rata OEE, OperatingRatio, Ability, Quality per machine per bulan
        [HttpGet("efficiency")]
        public IActionResult GetMachineEfficiency([FromQuery] int month, [FromQuery] int year)
        {
            if (month < 1 || month > 12)
                return BadRequest(new { error = "Bulan tidak valid (1–12)." });
            if (year < 2000 || year > 2100)
                return BadRequest(new { error = "Tahun tidak valid." });

            var result = new List<object>();
            try
            {
                using var conn = OpenConn();
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
            catch (SqlException sqlEx) { return StatusCode(500, new { error = $"Database error: {sqlEx.Message}" }); }
            catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
        }

        // ─── GET /api/machine/list ─────────────────────────────────────────────────
        // Daftar machine yang pernah ada di MachineEfficiency
        [HttpGet("list")]
        public IActionResult GetMachineList()
        {
            var result = new List<object>();
            try
            {
                using var conn = OpenConn();
                var query = @"
                    SELECT DISTINCT [MachineName]
                    FROM [dbo].[MachineEfficiency]
                    WHERE [MachineName] IS NOT NULL AND [MachineName] != ''
                    ORDER BY [MachineName]";

                using var cmd = new SqlCommand(query, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Add(new { machineName = reader["MachineName"]?.ToString() ?? "-" });

                return Ok(result);
            }
            catch (SqlException sqlEx) { return StatusCode(500, new { error = $"Database error: {sqlEx.Message}" }); }
            catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
        }

        // ─── GET /api/machine/detail?machineName=X&month=Y&year=Z ─────────────────
        // Detail harian 1 machine: OEE + breakdown loss per shift
        [HttpGet("detail")]
        public IActionResult GetMachineDetail(
            [FromQuery] string machineName,
            [FromQuery] int month,
            [FromQuery] int year)
        {
            if (string.IsNullOrEmpty(machineName))
                return BadRequest(new { error = "machineName wajib diisi." });
            if (month < 1 || month > 12)
                return BadRequest(new { error = "Bulan tidak valid." });

            var result = new List<object>();
            try
            {
                using var conn = OpenConn();
                var query = @"
                    SELECT
                        me.ID,
                        me.MachineName,
                        CONVERT(VARCHAR, me.[Date], 23)  AS [Date],
                        me.Shift,
                        me.OEE,
                        me.OperatingRatio,
                        me.Ability,
                        me.Quality,
                        me.Achievement,
                        me.WorkingTime,
                        me.PlanQty,
                        me.GoodProductionQty,
                        me.DefectQty,
                        mel.LossCategory,
                        mel.LossGroup,
                        mel.LossMinutes
                    FROM [dbo].[MachineEfficiency] me
                    LEFT JOIN [dbo].[MachineEfficiencyLoss] mel ON mel.EfficiencyID = me.ID
                    WHERE me.MachineName = @MachineName
                      AND MONTH(me.[Date]) = @Month
                      AND YEAR(me.[Date])  = @Year
                    ORDER BY me.[Date], me.Shift, mel.LossGroup, mel.LossCategory";

                using var cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@MachineName", machineName);
                cmd.Parameters.AddWithValue("@Month", month);
                cmd.Parameters.AddWithValue("@Year", year);

                // Group by header ID, flatten loss items
                var headers = new Dictionary<int, dynamic>();
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    int id = Convert.ToInt32(reader["ID"]);
                    if (!headers.ContainsKey(id))
                    {
                        headers[id] = new
                        {
                            id = id,
                            machineName = reader["MachineName"]?.ToString() ?? "-",
                            date = reader["Date"]?.ToString() ?? "-",
                            shift = reader["Shift"]?.ToString() ?? "-",
                            oee = reader["OEE"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["OEE"]),
                            operatingRatio = reader["OperatingRatio"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["OperatingRatio"]),
                            ability = reader["Ability"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["Ability"]),
                            quality = reader["Quality"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["Quality"]),
                            achievement = reader["Achievement"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["Achievement"]),
                            workingTime = reader["WorkingTime"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["WorkingTime"]),
                            planQty = reader["PlanQty"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["PlanQty"]),
                            goodProductionQty = reader["GoodProductionQty"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["GoodProductionQty"]),
                            defectQty = reader["DefectQty"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["DefectQty"]),
                            lossItems = new List<object>()
                        };
                    }
                    if (reader["LossCategory"] != DBNull.Value)
                    {
                        ((List<object>)headers[id].lossItems).Add(new
                        {
                            lossCategory = reader["LossCategory"]?.ToString() ?? "",
                            lossGroup = reader["LossGroup"]?.ToString() ?? "",
                            lossMinutes = reader["LossMinutes"] == DBNull.Value ? (double?)null : Convert.ToDouble(reader["LossMinutes"])
                        });
                    }
                }

                return Ok(headers.Values.ToList());
            }
            catch (SqlException sqlEx) { return StatusCode(500, new { error = $"Database error: {sqlEx.Message}" }); }
            catch (Exception ex) { return StatusCode(500, new { error = ex.Message }); }
        }
    }
}