using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using ClosedXML.Excel;
using MonitoringSystem.Data;
using System.Text.Json;
using MonitoringSystem.Models;
using OfficeOpenXml;
using System.Globalization;

namespace MonitoringSystem.Pages.LossTime
{
    public class IndexModel : PageModel
    {
        private readonly ApplicationDbContext _context;

        public IndexModel(ApplicationDbContext context, IConfiguration configuration, IWebHostEnvironment webHostEnvironment)
        {
            _context = context;
            _configuration = configuration;
            connectionString = _configuration.GetConnectionString("DefaultConnection") ?? "";
            machineConnectionString = _configuration.GetConnectionString("MachineConnection") ?? "";
            _webHostEnvironment = webHostEnvironment;
        }

        private string connectionString;
        private string machineConnectionString;
        private readonly IConfiguration _configuration;
        private readonly IWebHostEnvironment _webHostEnvironment;

        public List<LossTimeRecord> LossTimeData { get; set; } = new List<LossTimeRecord>();
        public int TotalDuration { get; set; }
        public int CurrentPage { get; set; } = 1;
        public int PageSize { get; set; } = 10;
        public int TotalPages => (int)Math.Ceiling((double)TotalRecords / PageSize);
        public int TotalRecords { get; set; }
        public bool HasDataToDisplay => TotalRecords > 0;

        [BindProperty]
        public DateTime StartSelectedDate { get; set; } = DateTime.Today;

        [BindProperty]
        public DateTime EndSelectedDate { get; set; } = DateTime.Today;

        [BindProperty]
        public int SelectedMonth { get; set; } = DateTime.Today.Month;
        [BindProperty]
        public int SelectedYear { get; set; } = DateTime.Today.Year;
        [BindProperty]
        public int TargetYear { get; set; } = DateTime.Today.Year;
        [BindProperty]
        public int TargetMonth { get; set; } = DateTime.Today.Month;

        [BindProperty]
        public string MachineLine { get; set; } = "All";

        [BindProperty]
        public string SelectedSource { get; set; } = "Assembly";

        [BindProperty]
        public string SelectedMachineName { get; set; } = "All";

        [BindProperty]
        public List<string> SelectedShifts { get; set; } = new List<string> { "1", "2", "3" };

        [BindProperty]
        public int SelectedPageSize { get; set; } = 10;

        [BindProperty]
        public string AdditionalBreakTime1Start { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime1End { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime2Start { get; set; } = "";
        [BindProperty]
        public string AdditionalBreakTime2End { get; set; } = "";
        [BindProperty]
        public IFormFile UploadedExcel { get; set; }
        [BindProperty]
        public string UploadMachineLine { get; set; }

        public bool IsFiltering { get; set; } = false;
        public Dictionary<string, int> CategorySummary { get; set; } = new Dictionary<string, int>();
        public string ChartDataJson { get; set; }
        public string DailyChartDataJson { get; set; }

        public List<string> MachineNameList { get; set; } = new List<string>();

        public List<string> AllCategories { get; set; } = new List<string>
        {
            "Model Changing Loss",
            "Material Shortage External",
            "Man Power Adjustment",
            "Material Shortage Internal",
            "Material Shortage Inhouse",
            "Quality Trouble",
            "Machine & Tools Trouble",
            "Rework",
            "Morning Assembly",
            "Reason Not Fill"
        };

        public Dictionary<string, string> CategoryAbbreviations = new()
        {
            { "Model Changing Loss", "Mdl Change" },
            { "Material Shortage External", "Mtrl Shortage Ex" },
            { "Man Power Adjustment", "MP Adjust" },
            { "Material Shortage Internal", "Mtrl Shortage Int" },
            { "Material Shortage Inhouse", "Mtrl Shortage Inhs" },
            { "Quality Trouble", "Quality" },
            { "Machine & Tools Trouble", "MC Trouble" },
            { "Rework", "Rework" },
            { "Morning Assembly", "Morning Assy" },
            { "Reason Not Fill", "Reason NF" }
        };

        private readonly Dictionary<string, string> CategoryColors = new Dictionary<string, string>
        {
            { "Model Changing Loss", "#FF6384" },
            { "Material Shortage External", "#36A2EB" },
            { "Man Power Adjustment", "#FFCE56" },
            { "Material Shortage Internal", "#4BC0C0" },
            { "Material Shortage Inhouse", "#9966FF" },
            { "Quality Trouble", "#FF9F40" },
            { "Machine & Tools Trouble", "#C9CBCF" },
            { "Rework", "#FF9F80" },
            { "Morning Assembly", "#198754" },
            { "Reason Not Fill", "#77DD77" }
        };

        private readonly List<(TimeSpan Start, TimeSpan End)> FixedBreakTimes = new List<(TimeSpan, TimeSpan)>
        {
            (new TimeSpan(7, 0, 0), new TimeSpan(7, 5, 0)),
            (new TimeSpan(9, 30, 0), new TimeSpan(9, 35, 0)),
            (new TimeSpan(15, 30, 0), new TimeSpan(15, 35, 0)),
            (new TimeSpan(18, 15, 0), new TimeSpan(18, 45, 0))
        };

        public void OnGet(int pageNumber = 1, int pageSize = 10)
        {
            CurrentPage = pageNumber;
            PageSize = pageSize;
            SelectedPageSize = pageSize;
            SetDatesFromMonthYear();
            LoadBreakTimeForToday();
            LoadMachineNameList();
            LoadData();
        }

        public void SetDatesFromMonthYear()
        {
            StartSelectedDate = new DateTime(SelectedYear, SelectedMonth, 1);
            EndSelectedDate = StartSelectedDate.AddMonths(1).AddDays(-1);
        }

        public IActionResult OnPostFilter()
        {
            _cachedBreakTimes = null;
            CurrentPage = 1;
            PageSize = SelectedPageSize;
            SetDatesFromMonthYear();
            if (SelectedShifts == null || !SelectedShifts.Any())
                SelectedShifts = new List<string> { "1", "2", "3" };
            LoadBreakTimeForToday();
            LoadMachineNameList();
            IsFiltering = true;
            LoadData();
            return Page();
        }

        public IActionResult OnPostChangePage(int pageNumber, int pageSize, int selectedMonth, int selectedYear,
            string machineLine, List<string> selectedShifts,
            string additionalBreakTime1Start, string additionalBreakTime1End,
            string additionalBreakTime2Start, string additionalBreakTime2End,
            string selectedSource = "Assembly", string selectedMachineName = "All")
        {
            CurrentPage = pageNumber;
            PageSize = pageSize;
            SelectedMonth = selectedMonth;
            SelectedYear = selectedYear;
            SetDatesFromMonthYear();
            MachineLine = machineLine;
            SelectedShifts = selectedShifts ?? new List<string> { "1", "2", "3" };
            AdditionalBreakTime1Start = additionalBreakTime1Start;
            AdditionalBreakTime1End = additionalBreakTime1End;
            AdditionalBreakTime2Start = additionalBreakTime2Start;
            AdditionalBreakTime2End = additionalBreakTime2End;
            SelectedSource = selectedSource;
            SelectedMachineName = selectedMachineName;
            LoadMachineNameList();
            LoadData();
            return Page();
        }

        public IActionResult OnPostReset()
        {
            _cachedBreakTimes = null;
            ModelState.Clear();
            SelectedMonth = DateTime.Today.Month;
            SelectedYear = DateTime.Today.Year;
            SetDatesFromMonthYear();
            MachineLine = "All";
            SelectedShifts = new List<string> { "1", "2", "3" };
            SelectedPageSize = 10;
            PageSize = 10;
            IsFiltering = false;
            CurrentPage = 1;
            SelectedSource = "Assembly";
            SelectedMachineName = "All";
            LoadBreakTimeForToday();
            LoadMachineNameList();
            LoadData();
            return Page();
        }

        private void LoadMachineNameList()
        {
            MachineNameList.Clear();
            try
            {
                if (string.IsNullOrEmpty(machineConnectionString))
                {
                    Console.WriteLine("❌ MachineConnection string is empty!");
                    return;
                }

                using var conn = new SqlConnection(machineConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(
                    "SELECT DISTINCT MachineName FROM [dbo].[MachineEfficiency] WHERE MachineName IS NOT NULL AND MachineName != '' ORDER BY MachineName",
                    conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    MachineNameList.Add(reader.GetString(0));

                Console.WriteLine($"✅ Loaded {MachineNameList.Count} machines");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ LoadMachineNameList error: {ex.Message}");
            }
        }

        private void LoadBreakTimeForToday()
        {
            var today = DateTime.Today;
            var latestBreakTime = _context.AdditionalBreakTimes.Where(bt => bt.Date == today).OrderByDescending(bt => bt.CreatedAt).FirstOrDefault();
            if (latestBreakTime != null)
            {
                AdditionalBreakTime1Start = latestBreakTime.BreakTime1Start?.ToString(@"hh\:mm");
                AdditionalBreakTime1End = latestBreakTime.BreakTime1End?.ToString(@"hh\:mm");
                AdditionalBreakTime2Start = latestBreakTime.BreakTime2Start?.ToString(@"hh\:mm");
                AdditionalBreakTime2End = latestBreakTime.BreakTime2End?.ToString(@"hh\:mm");
            }
        }

        private bool TryParseTimeSpan(string timeString, out TimeSpan result)
        {
            string[] formats = { "HH:mm", "H:mm", "HH:mm:ss", "H:mm:ss" };
            if (TimeSpan.TryParseExact(timeString, formats, null, out result)) return true;
            if (DateTime.TryParse(timeString, out DateTime dateTime)) { result = dateTime.TimeOfDay; return true; }
            result = TimeSpan.Zero; return false;
        }

        private bool IsInBreakTime(TimeSpan startTime, TimeSpan endTime, List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            foreach (var (breakStart, breakEnd) in breakTimes)
                if ((startTime >= breakStart && startTime <= breakEnd) || (endTime >= breakStart && endTime <= breakEnd) || (startTime <= breakStart && endTime >= breakEnd)) return true;
            return false;
        }

        private void LoadData()
        {
            if (SelectedSource == "Machine")
                LoadDataFromMachine();
            else
                LoadDataFromAssembly();
        }

        // ✅ FIX: Cek LossTimeActuals dulu, kalau ada pakai itu, kalau tidak fallback ke AssemblyLossTime
        private void LoadDataFromAssembly()
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();
            var breakTimes = GetAllBreakTimes();

            DateTime prevMonthDate = StartSelectedDate.AddMonths(-1);
            DateTime lastMonthStart = new DateTime(prevMonthDate.Year, prevMonthDate.Month, 1);
            DateTime lastMonthEnd = lastMonthStart.AddMonths(1).AddDays(-1);

            List<LossTimeRecord> currentRecords;
            List<LossTimeRecord> lastMonthRecords;

            // ✅ Cek apakah ada data Actuals untuk bulan/tahun/line yang dipilih
            if (HasActualsData())
            {
                Console.WriteLine("✅ Menggunakan data dari LossTimeActuals + AssemblyLossTime (merge)");

                // Ambil data dari Actuals (tanggal yang sudah di-upload Excel)
                var actualsRecords = GetActualsAsLossRecords();

                // Cari hari mana saja yang sudah ada di Actuals
                var daysInActuals = actualsRecords.Select(r => r.Date.Day).Distinct().ToHashSet();
                Console.WriteLine($"[Merge] Days covered by Actuals: {string.Join(",", daysInActuals)}");

                // Ambil AssemblyLossTime hanya untuk hari yang TIDAK ada di Actuals
                var allAssembly = GetCombinedRecords(
                    StartSelectedDate, EndSelectedDate,
                    StartSelectedDate, EndSelectedDate, breakTimes);

                var assemblyRecords = allAssembly
                    .Where(r => r.Date >= StartSelectedDate && r.Date <= EndSelectedDate)
                    .Where(r => !daysInActuals.Contains(r.Date.Day)) // ← skip hari yang sudah ada di Actuals
                    .ToList();

                Console.WriteLine($"[Merge] Actuals: {actualsRecords.Count} records, Assembly gap-fill: {assemblyRecords.Count} records");

                // Gabungkan: Actuals + Assembly (hari yang belum ada di Actuals)
                currentRecords = actualsRecords.Concat(assemblyRecords).ToList();

                // Last month tetap dari AssemblyLossTime untuk perbandingan chart
                var allLastMonth = GetCombinedRecords(lastMonthStart, lastMonthEnd, lastMonthStart, lastMonthEnd, breakTimes);
                lastMonthRecords = allLastMonth.Where(r => r.Date >= lastMonthStart && r.Date <= lastMonthEnd).ToList();
            }
            else
            {
                Console.WriteLine("ℹ️ Tidak ada Actuals, menggunakan AssemblyLossTime");
                var allRecords = GetCombinedRecords(lastMonthStart, lastMonthEnd, StartSelectedDate, EndSelectedDate, breakTimes);
                lastMonthRecords = allRecords.Where(r => r.Date >= lastMonthStart && r.Date <= lastMonthEnd).ToList();
                currentRecords = allRecords.Where(r => r.Date >= StartSelectedDate && r.Date <= EndSelectedDate).ToList();
            }

            PrepareSummaryChartData(currentRecords, lastMonthRecords);
            PrepareDailyChartData(currentRecords);

            // Paginate in-memory
            TotalRecords = currentRecords.Count;
            EnsureValidCurrentPage();
            LossTimeData = currentRecords
                .OrderByDescending(r => r.Date)
                .ThenBy(r => r.Location)
                .Skip((CurrentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            Console.WriteLine($"✅ LoadDataFromAssembly: {currentRecords.Count} records in {sw.ElapsedMilliseconds}ms");
        }

        // ✅ NEW: Cek apakah ada data Actuals untuk filter yang aktif
        private bool HasActualsData()
        {
            var query = _context.LossTimeActuals
                .Where(x => x.Month == SelectedMonth && x.Year == SelectedYear);

            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                query = query.Where(x => x.MachineLine == MachineLine);

            return query.Any();
        }

        // ✅ NEW: Konversi LossTimeActuals → LossTimeRecord untuk ditampilkan
        private List<LossTimeRecord> GetActualsAsLossRecords()
        {
            var results = new List<LossTimeRecord>();

            var query = _context.LossTimeActuals
                .Where(x => x.Month == SelectedMonth && x.Year == SelectedYear && x.Minutes > 0);

            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                query = query.Where(x => x.MachineLine == MachineLine);

            // Filter shift jika tidak semua shift dipilih
            var actuals = query.ToList();

            foreach (var actual in actuals)
            {
                // Validasi tanggal
                if (actual.Day < 1 || actual.Day > DateTime.DaysInMonth(actual.Year, actual.Month))
                    continue;

                results.Add(new LossTimeRecord
                {
                    Date = new DateTime(actual.Year, actual.Month, actual.Day),
                    Category = actual.Category,
                    Duration = (int)(actual.Minutes * 60), // menit → detik
                    Location = actual.MachineLine,
                    Shift = "1", // Actuals tidak punya info shift
                    LossTime = actual.Category,
                    Start = TimeSpan.Zero,
                    End = TimeSpan.Zero,
                    DetailedReason = $"Actual: {actual.Minutes} min"
                });
            }

            Console.WriteLine($"✅ GetActualsAsLossRecords: {results.Count} records");
            return results;
        }

        private void LoadDataFromMachine()
        {
            var sw = System.Diagnostics.Stopwatch.StartNew();

            DateTime prevMonthDate = StartSelectedDate.AddMonths(-1);
            DateTime lastMonthStart = new DateTime(prevMonthDate.Year, prevMonthDate.Month, 1);
            DateTime lastMonthEnd = lastMonthStart.AddMonths(1).AddDays(-1);

            var currentRecords = GetMachineRecords(StartSelectedDate, EndSelectedDate);
            var lastMonthRecords = GetMachineRecords(lastMonthStart, lastMonthEnd);

            Console.WriteLine($"🔍 Machine Records - Current: {currentRecords.Count}, LastMonth: {lastMonthRecords.Count}");

            PrepareSummaryChartData(currentRecords, lastMonthRecords);
            PrepareDailyChartData(currentRecords);

            TotalRecords = currentRecords.Count;
            EnsureValidCurrentPage();
            LossTimeData = currentRecords
                .OrderByDescending(r => r.Date)
                .Skip((CurrentPage - 1) * PageSize)
                .Take(PageSize)
                .ToList();

            Console.WriteLine($"✅ LoadDataFromMachine: {LossTimeData.Count} records in {sw.ElapsedMilliseconds}ms");
        }

        private List<LossTimeRecord> GetMachineRecords(DateTime start, DateTime end)
        {
            var records = new List<LossTimeRecord>();

            if (string.IsNullOrEmpty(machineConnectionString))
            {
                Console.WriteLine("❌ MachineConnection string is empty!");
                return records;
            }

            string machineFilter = (!string.IsNullOrEmpty(SelectedMachineName) && SelectedMachineName != "All")
                ? "AND me.MachineName = @MachineName" : "";

            string shiftFilter = "";
            if (SelectedShifts != null && SelectedShifts.Any() && SelectedShifts.Count < 3)
            {
                var shiftConds = new List<string>();
                if (SelectedShifts.Contains("1")) shiftConds.Add("me.Shift IN ('Shift 1','1')");
                if (SelectedShifts.Contains("2")) shiftConds.Add("me.Shift IN ('Shift 2','2')");
                if (SelectedShifts.Contains("3")) shiftConds.Add("me.Shift IN ('Shift 3','3')");
                if (shiftConds.Any()) shiftFilter = " AND (" + string.Join(" OR ", shiftConds) + ")";
            }

            string sql = $@"
                SELECT
                    me.MachineName,
                    me.[Date],
                    me.Shift,
                    mel.LossCategory,
                    mel.LossGroup,
                    mel.LossMinutes
                FROM [dbo].[MachineEfficiency] me
                INNER JOIN [dbo].[MachineEfficiencyLoss] mel
                    ON mel.EfficiencyID = me.ID
                WHERE CAST(me.[Date] AS DATE) >= @StartDate
                  AND CAST(me.[Date] AS DATE) <= @EndDate
                  AND mel.LossMinutes > 0
                  {machineFilter}
                  {shiftFilter}
                ORDER BY me.[Date] DESC, me.MachineName";

            try
            {
                using var conn = new SqlConnection(machineConnectionString);
                conn.Open();
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@StartDate", start.Date);
                cmd.Parameters.AddWithValue("@EndDate", end.Date);
                if (!string.IsNullOrEmpty(SelectedMachineName) && SelectedMachineName != "All")
                    cmd.Parameters.AddWithValue("@MachineName", SelectedMachineName);

                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    var date = reader.GetDateTime(reader.GetOrdinal("Date"));
                    var machineName = reader.IsDBNull(reader.GetOrdinal("MachineName")) ? "" : reader.GetString(reader.GetOrdinal("MachineName"));
                    var shiftRaw = reader.IsDBNull(reader.GetOrdinal("Shift")) ? "" : reader.GetString(reader.GetOrdinal("Shift"));
                    var lossCategory = reader.IsDBNull(reader.GetOrdinal("LossCategory")) ? "" : reader.GetString(reader.GetOrdinal("LossCategory"));
                    var lossGroup = reader.IsDBNull(reader.GetOrdinal("LossGroup")) ? "" : reader.GetString(reader.GetOrdinal("LossGroup"));
                    var lossMinutes = reader.IsDBNull(reader.GetOrdinal("LossMinutes")) ? 0.0 : Convert.ToDouble(reader["LossMinutes"]);

                    string shiftNum = shiftRaw switch
                    {
                        "Shift 1" or "1" => "1",
                        "Shift 2" or "2" => "2",
                        "Shift 3" or "3" => "3",
                        "Non Shift" or "NS" => "NS",
                        _ => shiftRaw
                    };

                    string category = CategorizeReason(lossCategory);

                    records.Add(new LossTimeRecord
                    {
                        Date = date,
                        LossTime = lossCategory,
                        Start = TimeSpan.Zero,
                        End = TimeSpan.Zero,
                        Duration = (int)(lossMinutes * 60),
                        Location = machineName,
                        Shift = shiftNum,
                        Category = category,
                        DetailedReason = $"{lossGroup} | {lossCategory}: {lossMinutes} min"
                    });
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ GetMachineRecords error: {ex.Message}\n{ex.StackTrace}");
            }

            return records;
        }

        private List<LossTimeRecord> GetCombinedRecords(
            DateTime lastStart, DateTime lastEnd,
            DateTime currStart, DateTime currEnd,
            List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            var records = new List<LossTimeRecord>();
            string query = BuildQueryForDateRange();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@StartDate", lastStart.Date);
                    command.Parameters.AddWithValue("@EndDate", currEnd.Date);

                    bool isHistorical = currEnd.Date < DateTime.Today;
                    command.Parameters.AddWithValue("@IsHistorical", isHistorical ? 1 : 0);

                    if (isHistorical)
                        command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59));
                    else
                        command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);

                    if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                        command.Parameters.AddWithValue("@MachineLine", MachineLine);

                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (IsInBreakTime(startTime, endTime, breakTimes)) continue;

                            string reason = reader.IsDBNull(reader.GetOrdinal("Reason")) ?
                                string.Empty : reader.GetString(reader.GetOrdinal("Reason"));

                            records.Add(new LossTimeRecord
                            {
                                Date = reader.GetDateTime(reader.GetOrdinal("Date")),
                                LossTime = reason,
                                Start = startTime,
                                End = endTime,
                                Duration = reader.IsDBNull(reader.GetOrdinal("LossTime")) ?
                                    0 : reader.GetInt32(reader.GetOrdinal("LossTime")),
                                Location = reader.IsDBNull(reader.GetOrdinal("MachineCode")) ?
                                    string.Empty : reader.GetString(reader.GetOrdinal("MachineCode")),
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ?
                                    string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
                                Category = CategorizeReason(reason),
                                DetailedReason = reader.IsDBNull(reader.GetOrdinal("DetailedReason")) ?
                                    null : reader.GetString(reader.GetOrdinal("DetailedReason"))
                            });
                        }
                    }
                }
            }
            return records;
        }

        private string BuildQueryForDateRange()
        {
            string query = @"
        SELECT Date, Reason, DetailedReason, MachineCode,
               CAST(Time AS TIME) AS StartTime, 
               CAST(EndDateTime AS TIME) AS EndTime, 
               LossTime, 
               CASE 
                   WHEN CAST(Time AS TIME) >= '07:00:00' AND CAST(Time AS TIME) < '15:45:00' THEN '1'
                   WHEN CAST(Time AS TIME) >= '15:45:00' AND CAST(Time AS TIME) < '23:15:00' THEN '2'
                   ELSE '3' 
               END AS Shift
        FROM AssemblyLossTime 
        WHERE (
            (Date = @StartDate AND CAST(Time AS TIME) >= '07:00:00')
            OR
            (Date > @StartDate AND Date < @EndDate)
            OR
            (@IsHistorical = 1 AND Date = DATEADD(DAY, 1, @EndDate) AND CAST(Time AS TIME) < '07:00:00')
            OR
            (@IsHistorical = 0 AND Date = @EndDate AND CAST(Time AS TIME) <= @CurrentTime)
        )";

            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                query += " AND MachineCode = @MachineLine";

            return query;
        }

        private void PrepareSummaryChartData(List<LossTimeRecord> currentRecords, List<LossTimeRecord> lastMonthRecords)
        {
            try
            {
                var categoryStats = AllCategories.Select(cat => new
                {
                    Name = cat,
                    S1 = currentRecords.Where(r => r.Category == cat && r.Shift == "1").Sum(r => r.Duration),
                    S2 = currentRecords.Where(r => r.Category == cat && r.Shift == "2").Sum(r => r.Duration),
                    S3 = currentRecords.Where(r => r.Category == cat && r.Shift == "3").Sum(r => r.Duration),
                    TotalCurrent = currentRecords.Where(r => r.Category == cat).Sum(r => r.Duration),
                    TotalLast = lastMonthRecords.Where(r => r.Category == cat).Sum(r => r.Duration)
                }).ToList();

                var sortedStats = categoryStats.OrderByDescending(x => x.TotalCurrent).ToList();

                var chartData = new
                {
                    labels = sortedStats.Select(x => CategoryAbbreviations.ContainsKey(x.Name) ? CategoryAbbreviations[x.Name] : x.Name).ToArray(),
                    fullLabels = sortedStats.Select(x => x.Name).ToArray(),
                    shift1Data = sortedStats.Select(x => Math.Round(x.S1 / 60.0, 2)).ToArray(),
                    shift2Data = sortedStats.Select(x => Math.Round(x.S2 / 60.0, 2)).ToArray(),
                    shift3Data = sortedStats.Select(x => Math.Round(x.S3 / 60.0, 2)).ToArray(),
                    lastMonthData = sortedStats.Select(x => Math.Round(x.TotalLast / 60.0, 2)).ToArray()
                };

                var options = new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
                ChartDataJson = JsonSerializer.Serialize(chartData, options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in PrepareSummaryChartData: {ex.Message}");
                ChartDataJson = "{}";
            }
        }

        private void PrepareDailyChartData(List<LossTimeRecord> currentRecords)
        {
            try
            {
                int daysInMonth = DateTime.DaysInMonth(SelectedYear, SelectedMonth);
                var days = Enumerable.Range(1, daysInMonth).ToArray();

                var dailyGroups = currentRecords
                    .GroupBy(r => new { Day = r.Date.Day, r.Category })
                    .ToDictionary(g => g.Key, g => g.Sum(x => x.Duration));

                var datasets = AllCategories.Select(category => new
                {
                    label = category,
                    data = days.Select(day =>
                    {
                        var key = new { Day = day, Category = category };
                        return dailyGroups.ContainsKey(key) ? Math.Round(dailyGroups[key] / 60.0, 2) : 0;
                    }).ToArray(),
                    backgroundColor = CategoryColors.ContainsKey(category) ? CategoryColors[category] : "#cccccc",
                    stack = "DayStack"
                }).ToList();

                var dailyChartData = new { labels = days.Select(d => d.ToString()).ToArray(), datasets = datasets };
                var options = new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping };
                DailyChartDataJson = JsonSerializer.Serialize(dailyChartData, options);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"❌ Error in PrepareDailyChartData: {ex.Message}");
                DailyChartDataJson = "{}";
            }
        }

        private void LoadPaginatedData(List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            TotalRecords = GetTotalRecords(breakTimes);
            EnsureValidCurrentPage();
            string query = BuildQueryBase();
            query += " ORDER BY [Date] DESC, Time OFFSET @Offset ROWS FETCH NEXT @PageSize ROWS ONLY";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    AddQueryParameters(command, StartSelectedDate, EndSelectedDate);
                    command.Parameters.AddWithValue("@Offset", (CurrentPage - 1) * PageSize);
                    command.Parameters.AddWithValue("@PageSize", PageSize);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        LossTimeData.Clear();
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (IsInBreakTime(startTime, endTime, breakTimes)) continue;
                            string reason = reader.IsDBNull(reader.GetOrdinal("Reason")) ? string.Empty : reader.GetString(reader.GetOrdinal("Reason"));
                            LossTimeData.Add(new LossTimeRecord
                            {
                                Nomor = reader.IsDBNull(reader.GetOrdinal("Id")) ? 0 : reader.GetInt32(reader.GetOrdinal("Id")),
                                Date = reader.IsDBNull(reader.GetOrdinal("Date")) ? DateTime.MinValue : reader.GetDateTime(reader.GetOrdinal("Date")),
                                LossTime = reason,
                                Start = startTime,
                                End = endTime,
                                Duration = reader.IsDBNull(reader.GetOrdinal("LossTime")) ? 0 : reader.GetInt32(reader.GetOrdinal("LossTime")),
                                Location = reader.IsDBNull(reader.GetOrdinal("MachineCode")) ? string.Empty : reader.GetString(reader.GetOrdinal("MachineCode")),
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ? string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
                                Category = CategorizeReason(reason),
                                DetailedReason = reader.IsDBNull(reader.GetOrdinal("DetailedReason")) ? null : reader.GetString(reader.GetOrdinal("DetailedReason"))
                            });
                        }
                    }
                }
            }
        }

        private List<LossTimeRecord> GetLossTimeRecords(DateTime start, DateTime end, List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            var records = new List<LossTimeRecord>();
            string query = BuildQueryForDateRange();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@StartDate", start.Date);
                    command.Parameters.AddWithValue("@EndDate", end.Date);
                    bool isHistorical = end.Date < DateTime.Today;
                    command.Parameters.AddWithValue("@IsHistorical", isHistorical ? 1 : 0);
                    if (isHistorical)
                        command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59));
                    else
                        command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);
                    if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                        command.Parameters.AddWithValue("@MachineLine", MachineLine);
                    using (SqlDataReader reader = command.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            TimeSpan startTime = reader.GetTimeSpan(reader.GetOrdinal("StartTime"));
                            TimeSpan endTime = reader.GetTimeSpan(reader.GetOrdinal("EndTime"));
                            if (IsInBreakTime(startTime, endTime, breakTimes)) continue;
                            string reason = reader.IsDBNull(reader.GetOrdinal("Reason")) ? string.Empty : reader.GetString(reader.GetOrdinal("Reason"));
                            records.Add(new LossTimeRecord
                            {
                                Date = reader.GetDateTime(reader.GetOrdinal("Date")),
                                LossTime = reason,
                                Start = startTime,
                                End = endTime,
                                Duration = reader.IsDBNull(reader.GetOrdinal("LossTime")) ? 0 : reader.GetInt32(reader.GetOrdinal("LossTime")),
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ? string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
                                Category = CategorizeReason(reason)
                            });
                        }
                    }
                }
            }
            return records;
        }

        private void CalculateAllDataSummary(List<(TimeSpan Start, TimeSpan End)> breakTimes) => PrepareDailyChartData(GetLossTimeRecords(StartSelectedDate, EndSelectedDate, breakTimes));

        private string BuildQueryBase()
        {
            bool isHistorical = EndSelectedDate.Date < DateTime.Today;
            string query = @"
        SELECT Id, Date, Reason, DetailedReason, 
               CAST(Time AS TIME) AS StartTime, 
               CAST(EndDateTime AS TIME) AS EndTime, 
               LossTime, MachineCode, 
               CASE 
                   WHEN CAST(Time AS TIME) >= '07:00:00' AND CAST(Time AS TIME) < '15:45:00' THEN '1'
                   WHEN CAST(Time AS TIME) >= '15:45:00' AND CAST(Time AS TIME) < '23:15:00' THEN '2'
                   ELSE '3' 
               END AS Shift
        FROM AssemblyLossTime 
        WHERE (
            (Date = @StartDate AND CAST(Time AS TIME) >= '07:00:00')
            OR
            (Date > @StartDate AND Date < @EndDate)
            OR
            (@IsHistorical = 1 AND Date = DATEADD(DAY, 1, @EndDate) AND CAST(Time AS TIME) < '07:00:00')
            OR
            (@IsHistorical = 0 AND Date = @EndDate AND CAST(Time AS TIME) <= @CurrentTime)
        )";

            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                query += " AND MachineCode = @MachineLine";

            if (SelectedShifts != null && SelectedShifts.Any() && SelectedShifts.Count < 3)
            {
                var shiftConditions = new List<string>();
                if (SelectedShifts.Contains("1")) shiftConditions.Add("(CAST(Time AS TIME) >= '07:00:00' AND CAST(Time AS TIME) < '15:45:00')");
                if (SelectedShifts.Contains("2")) shiftConditions.Add("(CAST(Time AS TIME) >= '15:45:00' AND CAST(Time AS TIME) < '23:15:00')");
                if (SelectedShifts.Contains("3")) shiftConditions.Add("(CAST(Time AS TIME) >= '23:15:00' OR CAST(Time AS TIME) < '07:00:00')");
                if (shiftConditions.Any()) query += " AND (" + string.Join(" OR ", shiftConditions) + ")";
            }

            return query;
        }

        private void AddQueryParameters(SqlCommand command, DateTime start, DateTime end)
        {
            bool isHistorical = end.Date < DateTime.Today;
            command.Parameters.AddWithValue("@StartDate", start.Date);
            command.Parameters.AddWithValue("@EndDate", end.Date);
            command.Parameters.AddWithValue("@IsHistorical", isHistorical ? 1 : 0);
            if (isHistorical)
                command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59));
            else
                command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);
            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                command.Parameters.AddWithValue("@MachineLine", MachineLine);
        }

        private int GetTotalRecords(List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            string baseQuery = BuildQueryBase();
            foreach (var (start, end) in breakTimes)
                baseQuery += $" AND NOT (CAST(Time AS TIME) BETWEEN '{start}' AND '{end}')";
            string countQuery = $"SELECT COUNT(*) FROM ({baseQuery}) AS CountQuery";
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(countQuery, connection))
                {
                    AddQueryParameters(command, StartSelectedDate, EndSelectedDate);
                    return (int)command.ExecuteScalar();
                }
            }
        }

        private List<(TimeSpan Start, TimeSpan End)> _cachedBreakTimes = null;

        private List<(TimeSpan Start, TimeSpan End)> GetAllBreakTimes()
        {
            if (_cachedBreakTimes != null) return _cachedBreakTimes;
            var breakTimes = new List<(TimeSpan Start, TimeSpan End)>();
            breakTimes.AddRange(FixedBreakTimes);
            if (!string.IsNullOrEmpty(AdditionalBreakTime1Start) && !string.IsNullOrEmpty(AdditionalBreakTime1End))
                if (TryParseTimeSpan(AdditionalBreakTime1Start, out TimeSpan start1) && TryParseTimeSpan(AdditionalBreakTime1End, out TimeSpan end1))
                    breakTimes.Add((start1, end1));
            if (!string.IsNullOrEmpty(AdditionalBreakTime2Start) && !string.IsNullOrEmpty(AdditionalBreakTime2End))
                if (TryParseTimeSpan(AdditionalBreakTime2Start, out TimeSpan start2) && TryParseTimeSpan(AdditionalBreakTime2End, out TimeSpan end2))
                    breakTimes.Add((start2, end2));
            _cachedBreakTimes = breakTimes;
            return breakTimes;
        }

        private void EnsureValidCurrentPage()
        {
            if (TotalRecords == 0) { CurrentPage = 1; return; }
            int maxPages = (int)Math.Ceiling((double)TotalRecords / PageSize);
            if (CurrentPage > maxPages) CurrentPage = maxPages;
            else if (CurrentPage < 1) CurrentPage = 1;
        }

        // ✅ FIX: CategorizeReason dengan exact match lengkap + fuzzy fallback
        private string CategorizeReason(string reason)
        {
            var r = reason?.ToLower().Trim() ?? "";
            if (string.IsNullOrEmpty(r)) return "Reason Not Fill";

            // ✅ Exact match PascalCase (dari Machine source)
            switch (reason?.Trim())
            {
                case "ModelChangingLoss": return "Model Changing Loss";
                case "MaterialShortageExternal": return "Material Shortage External";
                case "MaterialShortageInternal": return "Material Shortage Internal";
                case "MaterialShortageInhouse": return "Material Shortage Inhouse";
                case "ManPowerAdjustment": return "Man Power Adjustment";
                case "QualityTrouble": return "Quality Trouble";
                case "FreeTalkingQC": return "Quality Trouble";
                case "GeneralAssembly": return "Morning Assembly";
                case "TrialRun": return "Reason Not Fill";
            }

            // ✅ Exact match nama lengkap (dari AssemblyLossTime dan LossTimeActuals)
            switch (reason?.Trim())
            {
                case "Model Changing Loss": return "Model Changing Loss";
                case "Mold Changing Loss": return "Model Changing Loss";
                case "Material Shortage External": return "Material Shortage External";
                case "Gawse - External Loss": return "Material Shortage External";
                case "Gawse - External Bodies": return "Material Shortage External";
                case "Material Shortage Internal": return "Material Shortage Internal";
                case "Material Shortage Inhouse": return "Material Shortage Inhouse";
                case "Man Power Adjustment": return "Man Power Adjustment";
                case "Quality Trouble": return "Quality Trouble";
                case "Machine & Tools Trouble": return "Machine & Tools Trouble";
                case "Set Repairing Loss": return "Machine & Tools Trouble"; // ← KONFIRMASI jika perlu ganti
                case "Rework": return "Rework";
                case "Morning Assembly": return "Morning Assembly";
                case "General Assembly": return "Morning Assembly";
            }

            // ✅ Fuzzy match sebagai fallback terakhir
            if (r.Contains("model changing") || r.Contains("model change") || r.Contains("modelchanging") || r.Contains("mold changing"))
                return "Model Changing Loss";
            if (r.Contains("inhouse") || r.Contains("in house") || r.Contains("in-house"))
                return "Material Shortage Inhouse";
            if (r.Contains("internal") || r.Contains("materialshortageinternal"))
                return "Material Shortage Internal";
            if (r.Contains("gawse") || r.Contains("external") || r.Contains("materialshortageexternal"))
                return "Material Shortage External";
            if (r.Contains("material shortage") || r.Contains("materialshortage"))
                return "Material Shortage Internal";
            if (r.Contains("man power") || r.Contains("manpower") || r.Contains("manpoweradjustment"))
                return "Man Power Adjustment";
            if (r.Contains("quality") || r.Contains("freetalkingqc") || r.Contains("qc"))
                return "Quality Trouble";
            if (r.Contains("machine") || r.Contains("tools") || r.Contains("breakdown") || r.Contains("set repair"))
                return "Machine & Tools Trouble";
            if (r.Contains("rework") || r.Contains("re-work"))
                return "Rework";
            if (r.Contains("morning") || r.Contains("assembly") || r.Contains("generalassembly") || r.Contains("briefing"))
                return "Morning Assembly";

            Console.WriteLine($"⚠️ Unmatched LossCategory: '{reason}' → Reason Not Fill");
            return "Reason Not Fill";
        }

        // ✅ FIX: NormalizeCategoryName dengan mapping lengkap sebelum simpan ke DB
        private string NormalizeCategoryName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "SKIP";

            switch (input.Trim().ToUpper())
            {
                // ✅ Working Loss — mapping ke AllCategories
                case "QUALITY TROUBLE":
                case "FREE TALKING/QC ACTIVITY":
                case "FREE TALKING/QC":
                    return "Quality Trouble";

                case "MODEL CHANGING LOSS":
                case "MOLD CHANGING LOSS":
                    return "Model Changing Loss";

                case "MATERIAL SHORTAGE EXTERNAL":
                case "GAWSE - EXTERNAL BODIES":
                case "GAWSE - EXTERNAL LOSS":
                    return "Material Shortage External";

                case "MACHINE & TOOLS TROUBLE":
                case "SET REPAIRING LOSS": // ← KONFIRMASI jika perlu ganti
                    return "Machine & Tools Trouble";

                case "MAN POWER ADJUSTMENT":
                    return "Man Power Adjustment";

                case "MATERIAL SHORTAGE INHOUSE":
                    return "Material Shortage Inhouse";

                case "MATERIAL SHORTAGE INTERNAL":
                    return "Material Shortage Internal";

                case "REWORK":
                    return "Rework";

                case "MORNING ASSEMBLY":
                case "GENERAL ASSEMBLY":
                    return "Morning Assembly";

                // ✅ Fixed Loss — tidak perlu masuk chart, di-skip
                case "BREAK TIME (AM/PM)":
                case "COMPANY ACTIVITY":
                case "CLEANING":
                case "STOCK OPNAME":
                case "MAINTENANCE":
                case "TRIAL RUN":
                case "TRAINING EDUCATION":
                case "NO PRODUCTION DAY":
                    return "SKIP";

                default:
                    Console.WriteLine($"⚠️ Unmapped category from Excel: '{input}' → SKIP");
                    return "SKIP";
            }
        }

        public int GetTotalDurationAllCategories() => CategorySummary.Values.Sum();
        public double SecondsToMinutes(int seconds) => Math.Round(seconds / 60.0, 2);
        public List<int> GetPageSizeOptions() => new List<int> { 10 };

        public IActionResult OnPostExportExcel()
        {
            LoadBreakTimeForToday();
            SetDatesFromMonthYear();

            List<LossTimeRecord> exportData;

            if (SelectedSource == "Machine")
            {
                exportData = GetMachineRecords(StartSelectedDate, EndSelectedDate)
                    .OrderByDescending(x => x.Date).ToList();
            }
            else
            {
                // Export: gunakan Actuals kalau ada, fallback ke AssemblyLossTime
                if (HasActualsData())
                    exportData = GetActualsAsLossRecords().OrderByDescending(x => x.Date).ToList();
                else
                {
                    var breakTimes = GetAllBreakTimes();
                    exportData = GetCombinedRecords(StartSelectedDate, StartSelectedDate, StartSelectedDate, EndSelectedDate, breakTimes)
                        .OrderByDescending(x => x.Date).ToList();
                }
            }

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Loss Time Data");
                worksheet.Cell(1, 1).Value = "No";
                worksheet.Cell(1, 2).Value = "Date";
                worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(1, 4).Value = "Start Time";
                worksheet.Cell(1, 5).Value = "End Time";
                worksheet.Cell(1, 6).Value = "Duration (Min)";
                worksheet.Cell(1, 7).Value = "Location";
                worksheet.Cell(1, 8).Value = "Shift";
                worksheet.Cell(1, 9).Value = "Detailed Reason";

                int row = 2; int index = 1;
                foreach (var item in exportData)
                {
                    worksheet.Cell(row, 1).Value = index++;
                    worksheet.Cell(row, 2).Value = item.Date;
                    worksheet.Cell(row, 3).Value = item.Category;
                    worksheet.Cell(row, 4).Value = item.Start.ToString(@"hh\:mm\:ss");
                    worksheet.Cell(row, 5).Value = item.End.ToString(@"hh\:mm\:ss");
                    worksheet.Cell(row, 6).Value = Math.Round(item.Duration / 60.0, 2);
                    worksheet.Cell(row, 7).Value = item.Location;
                    worksheet.Cell(row, 8).Value = item.Shift;
                    worksheet.Cell(row, 9).Value = item.DetailedReason;
                    row++;
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return File(stream.ToArray(),
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        $"LossTime_{SelectedSource}_{StartSelectedDate:yyyyMMdd}-{EndSelectedDate:yyyyMMdd}.xlsx");
                }
            }
        }

        public IActionResult OnGetDownloadTemplateActualLoss()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "ActualLossTime", "Template_LossTime_Actual.xlsx");
            if (!System.IO.File.Exists(filePath)) return NotFound("File template tidak ditemukan di server.");
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Template_LossTime_Actual.xlsx");
        }

        public async Task<IActionResult> OnPostImportExcelActualAsync()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            if (UploadedExcel == null || UploadedExcel.Length == 0)
            {
                TempData["Error"] = "File Excel belum dipilih.";
                return RedirectToPage(new { TargetYear, TargetMonth, MachineLine });
            }

            if (string.IsNullOrEmpty(UploadMachineLine) || UploadMachineLine == "All")
            {
                TempData["Error"] = "Pilih Machine Line spesifik sebelum upload.";
                return RedirectToPage(new { TargetYear, TargetMonth, MachineLine });
            }

            if (TargetYear == 0) TargetYear = DateTime.Today.Year;
            if (TargetMonth == 0) TargetMonth = DateTime.Today.Month;

            Console.WriteLine($"[ImportActual] TargetYear={TargetYear}, TargetMonth={TargetMonth}, UploadMachineLine={UploadMachineLine}");

            try
            {
                using (var stream = new MemoryStream())
                {
                    await UploadedExcel.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var sheet = package.Workbook.Worksheets[0];

                        if (sheet.Dimension == null)
                        {
                            TempData["Error"] = "Sheet Excel kosong.";
                            return RedirectToPage(new { TargetYear, TargetMonth, MachineLine });
                        }

                        int rowCount = sheet.Dimension.Rows;
                        Console.WriteLine($"[ImportActual] Total rows in sheet: {rowCount}");

                        var newActuals = new List<LossTimeActual>();

                        for (int row = 3; row <= rowCount; row++)
                        {
                            var catName = sheet.Cells[row, 2].Value?.ToString()?.Trim();
                            if (string.IsNullOrEmpty(catName) ||
                                catName.ToLower().Contains("loss (min)") ||
                                catName.ToLower().Contains("loss category")) continue;

                            // ✅ FIX: Gunakan NormalizeCategoryName dengan mapping lengkap
                            catName = NormalizeCategoryName(catName);

                            // ✅ FIX: Skip kategori Fixed Loss atau yang tidak dikenal
                            if (catName == "SKIP") continue;

                            Console.WriteLine($"[ImportActual] Row {row}, Category: {catName}");

                            for (int day = 1; day <= 31; day++)
                            {
                                int col = 2 + day;
                                var cellValue = sheet.Cells[row, col].Value;
                                if (cellValue != null && double.TryParse(cellValue.ToString(), out double actualMinutes))
                                {
                                    if (day <= DateTime.DaysInMonth(TargetYear, TargetMonth))
                                    {
                                        newActuals.Add(new LossTimeActual
                                        {
                                            Category = catName,
                                            MachineLine = this.UploadMachineLine,
                                            Day = day,
                                            Month = TargetMonth,
                                            Year = TargetYear,
                                            Minutes = actualMinutes,
                                            CreatedAt = DateTime.Now
                                        });
                                    }
                                }
                            }
                        }

                        Console.WriteLine($"[ImportActual] Total data parsed: {newActuals.Count}");

                        // ✅ BARU — hapus hanya hari yang ada di Excel
                        if (newActuals.Any())
                        {
                            var daysInExcel = newActuals.Select(x => x.Day).Distinct().ToList();
                            Console.WriteLine($"[ImportActual] Days in Excel: {string.Join(",", daysInExcel)}");

                            var dataToDelete = _context.LossTimeActuals
                                .Where(x => x.MachineLine == this.UploadMachineLine &&
                                            x.Month == TargetMonth &&
                                            x.Year == TargetYear &&
                                            daysInExcel.Contains(x.Day)); // ← hanya hapus hari yang ada di Excel

                            _context.LossTimeActuals.RemoveRange(dataToDelete);
                            _context.LossTimeActuals.AddRange(newActuals);
                            await _context.SaveChangesAsync();
                            TempData["Success"] = $"Berhasil import {newActuals.Count} data Actual untuk {UploadMachineLine} (Bulan {TargetMonth}/{TargetYear}).";
                        }
                        else
                        {
                            TempData["Error"] = "Tidak ada data valid yang ditemukan di file Excel. Pastikan format sesuai template (data mulai dari baris 3, kolom 2).";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"[ImportActual] ERROR: {ex.Message}");
                TempData["Error"] = "Gagal Import: " + ex.Message;
            }

            return RedirectToPage(new { TargetYear, TargetMonth, MachineLine });
        }
    }

    public class LossTimeRecord
    {
        public int Nomor { get; set; }
        public DateTime Date { get; set; }
        public string LossTime { get; set; }
        public TimeSpan Start { get; set; }
        public TimeSpan End { get; set; }
        public int Duration { get; set; }
        public string Location { get; set; }
        public string Shift { get; set; }
        public string Category { get; set; }
        public string DetailedReason { get; set; }
    }
}