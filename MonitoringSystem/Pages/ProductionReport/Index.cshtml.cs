using Azure;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Math;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;
using System.Globalization;
using System.Text;
using static System.Runtime.InteropServices.JavaScript.JSType;

namespace MonitoringSystem.Pages.ProductionReport
{
    public class IndexModel : PageModel
    {
        private readonly IWebHostEnvironment _webHostEnvironment;
        private readonly IConfiguration _configuration;
        private string? connectionString;
        [BindProperty] public IFormFile? UploadedFile { get; set; } // Tambah ?
        [BindProperty] public string? TargetMachine { get; set; } // Tambah ?

        public bool IsCurrentMonthView { get; private set; }
        public int DaysInMonth { get; private set; }
        public List<string> ChartLabels { get; private set; } = new List<string>();
        public List<decimal> NormalData { get; private set; } = new List<decimal>();
        public List<decimal> OvertimeData { get; private set; } = new List<decimal>();
        public List<int> OriginalPlanData { get; private set; } = new List<int>();
        public List<int> PlanData { get; private set; } = new List<int>();

        //public List<int> EstimateData { get; private set; } = new List<int>();
        public List<int> NoOfDirectWorkers { get; private set; } = new List<int>();
        public List<int> DailyWorkTime { get; private set; } = new List<int>();
        public List<int> OvertimeOperators { get; private set; } = new List<int>();
        public List<int> OvertimeMinutes { get; private set; } = new List<int>();
        public List<int> DailyLossTime { get; private set; } = new List<int>();
        public List<int> PlanOvertimeData { get; private set; } = new List<int>();

        private class DailyData
        {
            public int Day { get; set; }

            // Normal shift data
            public decimal Shift1_Unit { get; set; }
            public TimeSpan Shift1_EndTime { get; set; }
            public decimal Shift2_Unit { get; set; }
            public TimeSpan Shift2_EndTime { get; set; }
            public decimal Shift3_Unit { get; set; }
            public TimeSpan Shift3_EndTime { get; set; }
            public decimal NonShift_Unit { get; set; }
            public TimeSpan NonShift_EndTime { get; set; }


            // Overtime data (Mode Shift OFF)
            public decimal Overtime_Unit { get; set; } = 0;
            public TimeSpan Overtime_EndTime { get; set; } = TimeSpan.Zero;

            public int Plan { get; set; }
            public int PlanOvertime { get; set; } = 0;
            public int OriginalPlan { get; set; } = 0;
            public int NoOfOperator { get; set; } = 0;
            public int OtOperatorCount { get; set; } = 0;
            public TimeSpan LastOtTime { get; set; } = TimeSpan.Zero;
        }

        public class RestTime { public int Duration { get; set; } public TimeSpan StartTime { get; set; } public TimeSpan EndTime { get; set; } }

        public IndexModel(IWebHostEnvironment webHostEnvironment, IConfiguration configuration)
        {
            _webHostEnvironment = webHostEnvironment;
            _configuration = configuration;
            // Pindahkan ke sini supaya semua fungsi sudah punya koneksi saat dipanggil
            this.connectionString = _configuration.GetConnectionString("DefaultConnection");
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }
        [BindProperty(SupportsGet = true)] public int SelectedMonth { get; set; } = DateTime.Now.Month;
        [BindProperty(SupportsGet = true)] public int SelectedYear { get; set; } = DateTime.Now.Year;
        [BindProperty(SupportsGet = true)] public string MachineLine { get; set; } = "All";
        [BindProperty(SupportsGet = true)] public List<string> SelectedShifts { get; set; } = new List<string>();
        public void OnGet()
        {
            // Normalisasi: jika kosong atau ada "All", set ke "All" saja
            if (!SelectedShifts.Any() || SelectedShifts.Contains("All"))
            {
                SelectedShifts = new List<string> { "All" };
            }
            // Jika user somehow pilih "All" + shift lain, prioritaskan "All"
            else if (SelectedShifts.Count > 1 && SelectedShifts.Contains("All"))
            {
                SelectedShifts = new List<string> { "All" };
            }

            LoadChartData();
        }

        public IActionResult OnPost(string submitButton)
        {
            if (submitButton == "reset")
            {
                return RedirectToPage(new { SelectedYear = DateTime.Now.Year, SelectedMonth = DateTime.Now.Month, MachineLine = "All" });
            }
            return RedirectToPage(new
            {
                SelectedYear = this.SelectedYear,
                SelectedMonth = this.SelectedMonth,
                MachineLine = this.MachineLine,
                SelectedShifts = this.SelectedShifts
            });
        }

        
        [BindProperty]
        public int TargetMonth { get; set; }
        [BindProperty]
        public int TargetYear { get; set; }

        public IActionResult OnGetDownloadTemplate(string type)
        {
            if (string.IsNullOrEmpty(type) || (type.ToLower() != "cu" && type.ToLower() != "cs"))
            {
                return NotFound("Invalid template type.");
            }
            string wwwrootPath = _webHostEnvironment.WebRootPath;
            var templateFileName = $"template_{type.ToLower()}.xlsx";
            var templateFilePath = Path.Combine(wwwrootPath, "data", type.ToLower(), "plan", $"{type.ToLower()}_plan_template.xlsx");
            if (System.IO.File.Exists(templateFilePath))
            {
                return File(System.IO.File.ReadAllBytes(templateFilePath), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", templateFileName);
            }
            return NotFound($"Template file not found.");
        }

        private void LoadChartData()
        {
            this.connectionString = _configuration.GetConnectionString("DefaultConnection");
            var dailyLosses = GetDailyLossTimeTotals();
            bool isCurrentMonthView = (SelectedYear == DateTime.Now.Year && SelectedMonth == DateTime.Now.Month);
            this.IsCurrentMonthView = isCurrentMonthView;

            string dateFilter = isCurrentMonthView ? "AND CAST(SDate AS DATE) <= @TodayDate" : "";
            this.DaysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);
            var combinedData = Enumerable.Range(1, this.DaysInMonth).Select(day => new DailyData { Day = day }).ToList();
            string shiftSelectionSql = "";
            if (!SelectedShifts.Contains("All") && SelectedShifts.Any())
            {
                var shiftConditions = new List<string>();
                foreach (var shift in SelectedShifts)
                {
                    // ⭐ TAMBAHKAN HANDLER UNTUK "NS"
                    if (shift == "NS" || shift == "ns")
                    {
                        shiftConditions.Add($"ShiftMode = 'NON-SHIFT'");
                    }
                    else
                    {
                        shiftConditions.Add($"ShiftMode = 'SHIFT {shift}'");
                    }
                }
                shiftSelectionSql = $"WHERE ({string.Join(" OR ", shiftConditions)})";
            }
            // 1. Query Plan
            string planShiftFilter = "";
            if (!SelectedShifts.Contains("All") && SelectedShifts.Any())
            {
                var conditions = SelectedShifts.Select(s => $"pr.shift LIKE '%{s}%'");
                planShiftFilter = $"AND ({string.Join(" OR ", conditions)})";
            }

            string planSql = $@"
        SELECT DAY(pp.CurrentDate) as Day, 
               SUM(ISNULL(pr.Quantity, 0)) as TotalPlanQuantity, 
               SUM(ISNULL(pr.Overtime, 0)) as TotalPlanOvertime
        FROM ProductionPlan pp
        LEFT JOIN ProductionRecords pr ON pp.Id = pr.PlanId
        WHERE YEAR(pp.CurrentDate) = @SelectedYear AND MONTH(pp.CurrentDate) = @SelectedMonth
        {planShiftFilter}
        {(MachineLine != "All" ? "AND pr.MachineCode = @MachineLine" : "AND pr.MachineCode IN ('MCH1-01', 'MCH1-02')")}
        GROUP BY DAY(pp.CurrentDate)";

            
string actualSql = $@"
WITH ShiftData AS(
    SELECT
        CASE
            WHEN CAST(SDate AS TIME) < '07:00:00'
            THEN CAST(DATEADD(DAY, -1, SDate) AS DATE)
            ELSE CAST(SDate AS DATE)
        END AS ReportDate,
            SDate,
            TotalUnit,
            NoOfOperator,
        ShiftMode,
        CASE WHEN ShiftMode = 'OVERTIME' THEN 1 ELSE 0 END as IsOvertimeFlag
    FROM oeesn
    WHERE(
        (YEAR(SDate) = @SelectedYear AND MONTH(SDate) = @SelectedMonth
         AND CAST(SDate AS TIME) >= '07:00:00')
        OR
        (SDate >= DATEADD(DAY, 1, DATEFROMPARTS(@SelectedYear, @SelectedMonth, 1))
         AND SDate < DATEADD(MONTH, 1, DATEFROMPARTS(@SelectedYear, @SelectedMonth, 1))
         AND CAST(SDate AS TIME) < '07:00:00')
    )
      { dateFilter}
            { (MachineLine != "All" ? "AND MachineCode = @MachineLine" : "AND MachineCode IN ('MCH1-01', 'MCH1-02')")}
),
ShiftDataFiltered AS(
    SELECT* FROM ShiftData
    { shiftSelectionSql}
),
DailyAggregates AS(
    SELECT ReportDate,
        MAX(CASE WHEN ShiftMode = 'SHIFT 1' THEN TotalUnit END) as S1_Unit,
        MAX(CASE WHEN ShiftMode = 'SHIFT 1' THEN CAST(SDate AS TIME) END) as S1_Time,
        MAX(CASE WHEN ShiftMode = 'SHIFT 2' THEN TotalUnit END) as S2_Unit,
        MAX(CASE WHEN ShiftMode = 'SHIFT 2' THEN CAST(SDate AS TIME) END) as S2_Time,
        MAX(CASE WHEN ShiftMode = 'SHIFT 3' THEN TotalUnit END) as S3_Unit,
        MAX(CASE WHEN ShiftMode = 'SHIFT 3' THEN CAST(SDate AS TIME) END) as S3_Time,
        MAX(CASE WHEN ShiftMode = 'NON-SHIFT' THEN TotalUnit END) as NS_Unit,
        MAX(CASE WHEN ShiftMode = 'NON-SHIFT' THEN CAST(SDate AS TIME) END) as NS_Time,
        MAX(CASE WHEN ShiftMode = 'OVERTIME' THEN TotalUnit END) as OT_Unit,
        MAX(CASE WHEN ShiftMode = 'OVERTIME' THEN CAST(SDate AS TIME) END) as OT_Time,
        MAX(NoOfOperator) as MaxOp
    FROM ShiftDataFiltered
    GROUP BY ReportDate
)
SELECT DAY(ReportDate) as Day, *FROM DailyAggregates";
            try
            {
                using (var conn = new SqlConnection(this.connectionString))
                {
                    conn.Open();

                    // Ambil Data Plan
                    using (var planCmd = new SqlCommand(planSql, conn))
                    {
                        planCmd.Parameters.AddWithValue("@SelectedYear", SelectedYear);
                        planCmd.Parameters.AddWithValue("@SelectedMonth", SelectedMonth);
                        if (MachineLine != "All") planCmd.Parameters.AddWithValue("@MachineLine", MachineLine);

                        using (var reader = planCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var d = combinedData.FirstOrDefault(x => x.Day == (int)reader["Day"]);
                                if (d != null)
                                {
                                    d.Plan = Convert.ToInt32(reader["TotalPlanQuantity"]);
                                    d.PlanOvertime = Convert.ToInt32(reader["TotalPlanOvertime"]);
                                }
                            }
                        }
                    }

                    // Ambil Data Actual
                    using (var actualCmd = new SqlCommand(actualSql, conn))
                    {
                        actualCmd.Parameters.AddWithValue("@SelectedYear", SelectedYear);
                        actualCmd.Parameters.AddWithValue("@SelectedMonth", SelectedMonth);
                        if (isCurrentMonthView) actualCmd.Parameters.AddWithValue("@TodayDate", DateTime.Now.Date);
                        if (MachineLine != "All") actualCmd.Parameters.AddWithValue("@MachineLine", MachineLine);

                        using (var reader = actualCmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var d = combinedData.FirstOrDefault(x => x.Day == (int)reader["Day"]);
                                if (d != null)
                                {
                                    d.Shift1_Unit = reader["S1_Unit"] != DBNull.Value ? (decimal)reader["S1_Unit"] : 0;
                                    d.Shift1_EndTime = reader["S1_Time"] != DBNull.Value ? (TimeSpan)reader["S1_Time"] : TimeSpan.Zero;

                                    d.Shift2_Unit = reader["S2_Unit"] != DBNull.Value ? (decimal)reader["S2_Unit"] : 0;
                                    d.Shift2_EndTime = reader["S2_Time"] != DBNull.Value ? (TimeSpan)reader["S2_Time"] : TimeSpan.Zero;

                                    d.Shift3_Unit = reader["S3_Unit"] != DBNull.Value ? (decimal)reader["S3_Unit"] : 0;
                                    d.Shift3_EndTime = reader["S3_Time"] != DBNull.Value ? (TimeSpan)reader["S3_Time"] : TimeSpan.Zero;

                                    // ⭐ TAMBAH 2 BARIS INI
                                    d.NonShift_Unit = reader["NS_Unit"] != DBNull.Value ? (decimal)reader["NS_Unit"] : 0;
                                    d.NonShift_EndTime = reader["NS_Time"] != DBNull.Value ? (TimeSpan)reader["NS_Time"] : TimeSpan.Zero;

                                    d.Overtime_Unit = reader["OT_Unit"] != DBNull.Value ? (decimal)reader["OT_Unit"] : 0;
                                    d.Overtime_EndTime = reader["OT_Time"] != DBNull.Value ? (TimeSpan)reader["OT_Time"] : TimeSpan.Zero;

                                    d.NoOfOperator = reader["MaxOp"] != DBNull.Value ? (int)reader["MaxOp"] : 0;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error LoadChartData: " + ex.Message);
            }

            // 3. Hitung Kalkulasi Akhir
            foreach (var data in combinedData)
            {
                ChartLabels.Add(data.Day.ToString());
                PlanData.Add(data.Plan);
                PlanOvertimeData.Add(data.PlanOvertime);

                int totalOtMinutes = 0;

                // CASE 1: Mode Shift OFF (NS atau OVERTIME)
                if (data.Overtime_Unit > 0)
                {
                    TimeSpan workStart = new TimeSpan(7, 0, 0);
                    if (data.Overtime_EndTime > workStart)
                    {
                        totalOtMinutes = (int)(data.Overtime_EndTime - workStart).TotalMinutes;
                    }
                    else if (data.Overtime_EndTime < workStart)
                    {
                        int minutesToMidnight = (int)(new TimeSpan(24, 0, 0) - workStart).TotalMinutes;
                        int minutesAfterMidnight = (int)data.Overtime_EndTime.TotalMinutes;
                        totalOtMinutes = minutesToMidnight + minutesAfterMidnight;
                    }
                }
                // ⭐ TAMBAH CASE BARU: NON-SHIFT yang overtime (lewat jam 16:00)
                else if (data.NonShift_Unit > 0 && data.NonShift_EndTime > new TimeSpan(16, 0, 0))
                {
                    totalOtMinutes = (int)(data.NonShift_EndTime - new TimeSpan(16, 0, 0)).TotalMinutes;
                }
                // CASE 2: Mode Shift ON
                else
                {
                    if (data.Shift1_EndTime > new TimeSpan(15, 45, 0))
                    {
                        totalOtMinutes += (int)(data.Shift1_EndTime - new TimeSpan(15, 45, 0)).TotalMinutes;
                    }

                    if (data.Shift2_EndTime > new TimeSpan(23, 0, 0))
                    {
                        if (data.Shift2_EndTime <= new TimeSpan(23, 59, 59))
                        {
                            totalOtMinutes += (int)(data.Shift2_EndTime - new TimeSpan(23, 0, 0)).TotalMinutes;
                        }
                        else if (data.Shift2_EndTime < new TimeSpan(7, 0, 0))
                        {
                            int minutesToMidnight = 59;
                            int minutesAfterMidnight = (int)data.Shift2_EndTime.TotalMinutes;
                            totalOtMinutes += minutesToMidnight + minutesAfterMidnight;
                        }
                    }

                    if (data.Shift3_EndTime > new TimeSpan(7, 0, 0) &&
                        data.Shift3_EndTime < new TimeSpan(15, 0, 0))
                    {
                        totalOtMinutes += (int)(data.Shift3_EndTime - new TimeSpan(7, 0, 0)).TotalMinutes;
                    }
                }

                OvertimeMinutes.Add(totalOtMinutes);

                // ✅ FIX: Overtime operator hanya kalau benar-benar ada overtime
                int overtimeOpCount = (data.Overtime_Unit > 0 || totalOtMinutes > 0)
                    ? (data.Overtime_Unit > 0
                        ? data.NoOfOperator  // Murni OVERTIME mode
                        : data.NoOfOperator) // Shift biasa yang lembur
                    : 0;
                OvertimeOperators.Add(overtimeOpCount);
                // ✅ UNIT CALCULATION
                decimal normalUnits = 0;
                decimal overtimeUnits = 0;
                // ✅ FIX: NON-SHIFT = Normal, OVERTIME = Overtime
                if (data.Overtime_Unit > 0 && data.Shift1_Unit == 0 && data.Shift2_Unit == 0 && data.Shift3_Unit == 0 && data.NonShift_Unit == 0)
                {
                    // Murni OVERTIME saja
                    normalUnits = 0;
                    overtimeUnits = data.Overtime_Unit;
                }
                else
                {
                    // NON-SHIFT + SHIFT biasa = Normal (biru)
                    normalUnits = data.Shift1_Unit + data.Shift2_Unit + data.Shift3_Unit + data.NonShift_Unit;

                    // Jika ada OVERTIME di hari yang sama dengan shift normal
                    overtimeUnits = data.Overtime_Unit;
                }

                NormalData.Add(normalUnits);   // BIRU
                OvertimeData.Add(overtimeUnits); // HIJAUAU

                NoOfDirectWorkers.Add(data.NoOfOperator);

                // ✅ LOSS TIME
                dailyLosses.TryGetValue(data.Day, out int lossDurationSec);
                DailyLossTime.Add(lossDurationSec / 60);

                // ✅ TOTAL WORK TIME
                var dayType = DetermineTypeOfDay(new DateTime(SelectedYear, SelectedMonth, data.Day).DayOfWeek);
                int stdWorkingMinutes = (dayType == "FRIDAY") ? 435 : (dayType == "WEEKEND" ? 0 : 473);

                int baseWorkMinutes = (normalUnits > 0 || overtimeUnits > 0) ? stdWorkingMinutes : 0;
                DailyWorkTime.Add(baseWorkMinutes + totalOtMinutes);
            }
        }
        private readonly List<(TimeSpan Start, TimeSpan End)> RegularDayBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(12, 0, 0), new TimeSpan(12, 45, 0)),
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))
        };

        private readonly List<(TimeSpan Start, TimeSpan End)> FridayBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(11, 50, 0), new TimeSpan(13, 15, 0)),
            (new TimeSpan(14, 30, 0), new TimeSpan(14, 35, 0))
        };
        private bool IsInBreakTime(TimeSpan startTime, TimeSpan endTime, List<(TimeSpan Start, TimeSpan End)> breakTimes) { foreach (var (breakStart, breakEnd) in breakTimes) { if (startTime < breakEnd && endTime > breakStart) { return true; } } return false; }
        private Dictionary<int, int> GetDailyLossTimeTotals()
        {
            var dailyTotals = new Dictionary<int, int>();

            // Ganti logika CASE WHEN yang ribet tadi dengan filter jam sederhana
            string shiftFilterSql = "";
            if (SelectedShifts.Any() && !SelectedShifts.Contains("All"))
            {
                var hours = new List<string>();
                foreach (var shift in SelectedShifts)
                {
                    if (shift == "1") hours.Add("(DATEPART(HOUR, Time) >= 7 AND DATEPART(HOUR, Time) < 16)");
                    if (shift == "2") hours.Add("((DATEPART(HOUR, Time) >= 16 AND DATEPART(HOUR, Time) < 23) OR (DATEPART(HOUR, Time) = 23 AND DATEPART(MINUTE, Time) <= 15))");
                    if (shift == "3") hours.Add("(DATEPART(HOUR, Time) >= 23 OR DATEPART(HOUR, Time) < 7)");
                }
                shiftFilterSql = $"AND ({string.Join(" OR ", hours)})";
            }

            string lossTimeMachineFilter = (MachineLine == "All")
                ? "AND MachineCode IN ('MCH1-01', 'MCH1-02')"
                : "AND MachineCode = @Machine";

            string query = $@"
        SELECT 
            CAST(Date AS DATE) as FullDate,
            CAST(Time AS TIME) as StartTime, 
            CAST(EndDateTime AS TIME) as EndTime, 
            LossTime as Duration
        FROM AssemblyLossTime
        WHERE YEAR(Date) = @Year 
          AND MONTH(Date) = @Month 
          {lossTimeMachineFilter}
          {shiftFilterSql}";

            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    using (var command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@Year", SelectedYear);
                        command.Parameters.AddWithValue("@Month", SelectedMonth);

                        if (MachineLine != "All")
                        {
                            command.Parameters.AddWithValue("@Machine", MachineLine);
                        }

                        connection.Open();
                        using (var reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                var fullDate = (DateTime)reader["FullDate"];
                                var startTime = (TimeSpan)reader["StartTime"];

                                // Shift 3 adjustment: jam 00:00 - 06:59 dianggap milik hari sebelumnya
                                if (startTime >= TimeSpan.Zero && startTime < new TimeSpan(7, 0, 0))
                                {
                                    fullDate = fullDate.AddDays(-1);
                                }

                                var day = fullDate.Day;
                                var endTime = (TimeSpan)reader["EndTime"];
                                var duration = Convert.ToInt32(reader["Duration"]);

                                var dayType = DetermineTypeOfDay(fullDate.DayOfWeek);
                                var breaksForThisDay = (dayType == "FRIDAY") ? this.FridayBreakTimes : this.RegularDayBreakTimes;

                                if (!IsInBreakTime(startTime, endTime, breaksForThisDay))
                                {
                                    if (!dailyTotals.ContainsKey(day))
                                    {
                                        dailyTotals[day] = 0;
                                    }
                                    dailyTotals[day] += duration;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error fetching loss time: {ex.Message}");
            }
            return dailyTotals;
        }
        //ini kayaknya ga kepakek hehe
        private List<(TimeSpan Start, TimeSpan End)> GetAdditionalBreakTimesForDate(DateTime date)
        {
            var additionalBreaks = new List<(TimeSpan, TimeSpan)>();
            try
            {
                using (var connection = new SqlConnection(this.connectionString))
                {
                    connection.Open();
                    string sql = @"
                SELECT TOP 1 BreakTime1Start, BreakTime1End, BreakTime2Start, BreakTime2End 
                FROM AdditionalBreakTimes 
                WHERE CAST(Date AS DATE) = @Date
                ORDER BY CreatedAt DESC";
                    using (var command = new SqlCommand(sql, connection))
                    {
                        command.Parameters.AddWithValue("@Date", date.Date);
                        using (var reader = command.ExecuteReader())
                        {
                            if (reader.Read())
                            {
                                if (!reader.IsDBNull(0) && !reader.IsDBNull(1))
                                    additionalBreaks.Add((reader.GetTimeSpan(0), reader.GetTimeSpan(1)));
                                if (!reader.IsDBNull(2) && !reader.IsDBNull(3))
                                    additionalBreaks.Add((reader.GetTimeSpan(2), reader.GetTimeSpan(3)));
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting additional breaks: {ex.Message}");
            }
            return additionalBreaks;
        }
        public int GetTotalRestTime(List<RestTime> listRestTime, TimeSpan StartTime, TimeSpan EndTime, TimeSpan CurrentTime) { int TotalRestTime = 0; bool isToday = (SelectedYear == DateTime.Now.Year && SelectedMonth == DateTime.Now.Month); TotalRestTime = listRestTime.Sum(rest => { if (isToday && rest.StartTime > CurrentTime) { return 0; } TimeSpan effectiveRestStart = rest.StartTime < StartTime ? StartTime : rest.StartTime; TimeSpan effectiveRestEnd = rest.EndTime > EndTime ? EndTime : rest.EndTime; if (isToday && effectiveRestEnd > CurrentTime) { effectiveRestEnd = CurrentTime; } return effectiveRestStart < effectiveRestEnd ? (int)(effectiveRestEnd - effectiveRestStart).TotalMinutes : 0; }); return TotalRestTime; }
        public string DetermineTypeOfDay(DayOfWeek day) { return day switch { DayOfWeek.Monday or DayOfWeek.Tuesday or DayOfWeek.Wednesday or DayOfWeek.Thursday => "REGULAR", DayOfWeek.Friday => "FRIDAY", DayOfWeek.Saturday or DayOfWeek.Sunday => "WEEKEND", _ => "REGULAR" }; }
        public List<RestTime> GetRestTime(string dayTipe) { List<RestTime> listRestTime = new List<RestTime>(); try { using (SqlConnection connection = new SqlConnection(this.connectionString)) { connection.Open(); string GetRestTime = @"SELECT Duration, StartTime, EndTime FROM RestTime WHERE DayType = @DayTipe"; using (SqlCommand command = new SqlCommand(GetRestTime, connection)) { command.Parameters.AddWithValue("@DayTipe", dayTipe); using (SqlDataReader dataReader = command.ExecuteReader()) { while (dataReader.Read()) { if (!dataReader.IsDBNull(0)) { listRestTime.Add(new RestTime { Duration = dataReader.GetInt32(0), StartTime = dataReader.GetTimeSpan(1), EndTime = dataReader.GetTimeSpan(2) }); } } } } } } catch (Exception ex) { Console.WriteLine("Exception GetRestTime: " + ex.ToString()); } return listRestTime; }
    }
}