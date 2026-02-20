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
			_webHostEnvironment = webHostEnvironment;
		}

		private string connectionString;
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
            LoadData();
        }

        public void SetDatesFromMonthYear()
        {
            StartSelectedDate = new DateTime(SelectedYear, SelectedMonth, 1);
            EndSelectedDate = StartSelectedDate.AddMonths(1).AddDays(-1);
        }

        public IActionResult OnPostFilter()
        {
            _cachedBreakTimes = null; // ✅ Clear cache
            CurrentPage = 1;
            PageSize = SelectedPageSize;
            SetDatesFromMonthYear();
            if (SelectedShifts == null || !SelectedShifts.Any())
                SelectedShifts = new List<string> { "1", "2", "3" };
            LoadBreakTimeForToday();
            IsFiltering = true;
            LoadData();
            return Page();
        }


        public IActionResult OnPostChangePage(int pageNumber, int pageSize, int selectedMonth, int selectedYear,
            string machineLine, List<string> selectedShifts,
            string additionalBreakTime1Start, string additionalBreakTime1End,
            string additionalBreakTime2Start, string additionalBreakTime2End)
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
            LoadData();
            return Page();
        }

        public IActionResult OnPostReset()
        {
            _cachedBreakTimes = null; // ✅ Clear cache
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
            LoadBreakTimeForToday();
            LoadData();
            return Page();
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

       /*private List<(TimeSpan Start, TimeSpan End)> GetAllBreakTimes()
        {
            var breakTimes = new List<(TimeSpan Start, TimeSpan End)>();
            breakTimes.AddRange(FixedBreakTimes);
            if (!string.IsNullOrEmpty(AdditionalBreakTime1Start) && !string.IsNullOrEmpty(AdditionalBreakTime1End))
                if (TryParseTimeSpan(AdditionalBreakTime1Start, out TimeSpan start1) && TryParseTimeSpan(AdditionalBreakTime1End, out TimeSpan end1)) breakTimes.Add((start1, end1));
            if (!string.IsNullOrEmpty(AdditionalBreakTime2Start) && !string.IsNullOrEmpty(AdditionalBreakTime2End))
                if (TryParseTimeSpan(AdditionalBreakTime2Start, out TimeSpan start2) && TryParseTimeSpan(AdditionalBreakTime2End, out TimeSpan end2)) breakTimes.Add((start2, end2));
            return breakTimes;
        } */

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
            var sw = System.Diagnostics.Stopwatch.StartNew();

            var breakTimes = GetAllBreakTimes();
            Console.WriteLine($"GetAllBreakTimes: {sw.ElapsedMilliseconds}ms");
            sw.Restart();

            DateTime prevMonthDate = StartSelectedDate.AddMonths(-1);
            DateTime lastMonthStart = new DateTime(prevMonthDate.Year, prevMonthDate.Month, 1);
            DateTime lastMonthEnd = lastMonthStart.AddMonths(1).AddDays(-1);

            var allRecords = GetCombinedRecords(
                lastMonthStart, lastMonthEnd,
                StartSelectedDate, EndSelectedDate,
                breakTimes
            );
            Console.WriteLine($"GetCombinedRecords: {sw.ElapsedMilliseconds}ms");
            sw.Restart();

            var lastMonthRecords = allRecords.Where(r => r.Date >= lastMonthStart && r.Date <= lastMonthEnd).ToList();
            var currentRecords = allRecords.Where(r => r.Date >= StartSelectedDate && r.Date <= EndSelectedDate).ToList();

            PrepareSummaryChartData(currentRecords, lastMonthRecords);
            Console.WriteLine($"PrepareSummaryChartData: {sw.ElapsedMilliseconds}ms");
            sw.Restart();

            PrepareDailyChartData(currentRecords);
            Console.WriteLine($"PrepareDailyChartData: {sw.ElapsedMilliseconds}ms");
            sw.Restart();

            LoadPaginatedData(breakTimes);
            Console.WriteLine($"LoadPaginatedData: {sw.ElapsedMilliseconds}ms");
            sw.Restart();

            Console.WriteLine($"TOTAL LoadData: {sw.ElapsedMilliseconds}ms");
        }

        // Method baru untuk gabung query
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

                    // ✅ Selalu set @CurrentTime (tidak boleh NULL)
                    if (isHistorical)
                    {
                        command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59)); // Dummy
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);
                    }

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
                    labels = sortedStats
                        .Select(x => CategoryAbbreviations.ContainsKey(x.Name)
                            ? CategoryAbbreviations[x.Name]
                            : x.Name)
                        .ToArray(),
                    fullLabels = sortedStats
                        .Select(x => x.Name)
                        .ToArray(),
                    shift1Data = sortedStats.Select(x => Math.Round(x.S1 / 60.0, 2)).ToArray(),
                    shift2Data = sortedStats.Select(x => Math.Round(x.S2 / 60.0, 2)).ToArray(),
                    shift3Data = sortedStats.Select(x => Math.Round(x.S3 / 60.0, 2)).ToArray(),
                    lastMonthData = sortedStats.Select(x => Math.Round(x.TotalLast / 60.0, 2)).ToArray()
                };

                // ✅ UBAH INI - Tambah options untuk tidak escape HTML
                var options = new JsonSerializerOptions
                {
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                ChartDataJson = JsonSerializer.Serialize(chartData, options);
                Console.WriteLine($"📊 ChartDataJson: {ChartDataJson}");
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

                var dailyChartData = new
                {
                    labels = days.Select(d => d.ToString()).ToArray(),
                    datasets = datasets
                };

                // ✅ UBAH INI - Tambah options untuk tidak escape HTML
                var options = new JsonSerializerOptions
                {
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
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
        // Tambahkan method ini (belum ada di code)
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
                    {
                        command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59));
                    }
                    else
                    {
                        command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);
                    }

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
                                Shift = reader.IsDBNull(reader.GetOrdinal("Shift")) ?
                                    string.Empty : reader.GetString(reader.GetOrdinal("Shift")),
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

                if (SelectedShifts.Contains("1"))
                    shiftConditions.Add("(CAST(Time AS TIME) >= '07:00:00' AND CAST(Time AS TIME) < '15:45:00')");

                if (SelectedShifts.Contains("2"))
                    shiftConditions.Add("(CAST(Time AS TIME) >= '15:45:00' AND CAST(Time AS TIME) < '23:15:00')");

                if (SelectedShifts.Contains("3"))
                    shiftConditions.Add("(CAST(Time AS TIME) >= '23:15:00' OR CAST(Time AS TIME) < '07:00:00')");

                if (shiftConditions.Any())
                    query += " AND (" + string.Join(" OR ", shiftConditions) + ")";
            }

            return query;
        }

        private void AddQueryParameters(SqlCommand command, DateTime start, DateTime end)
        {
            bool isHistorical = end.Date < DateTime.Today;

            command.Parameters.AddWithValue("@StartDate", start.Date);
            command.Parameters.AddWithValue("@EndDate", end.Date);
            command.Parameters.AddWithValue("@IsHistorical", isHistorical ? 1 : 0);

            // ✅ Selalu set @CurrentTime (tidak boleh NULL)
            if (isHistorical)
            {
                command.Parameters.AddWithValue("@CurrentTime", new TimeSpan(23, 59, 59)); // Dummy
            }
            else
            {
                command.Parameters.AddWithValue("@CurrentTime", DateTime.Now.TimeOfDay);
            }

            if (!string.IsNullOrEmpty(MachineLine) && MachineLine != "All")
                command.Parameters.AddWithValue("@MachineLine", MachineLine);
        }

        private int GetTotalRecords(List<(TimeSpan Start, TimeSpan End)> breakTimes)
        {
            string baseQuery = BuildQueryBase(); // ✅ Sudah benar sekarang

            // Tambah filter break time
            foreach (var (start, end) in breakTimes)
            {
                baseQuery += $" AND NOT (CAST(Time AS TIME) BETWEEN '{start}' AND '{end}')";
            }

            string countQuery = $"SELECT COUNT(*) FROM ({baseQuery}) AS CountQuery";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(countQuery, connection))
                {
                    AddQueryParameters(command, StartSelectedDate, EndSelectedDate); // ✅ Pakai method yang sudah di-update
                    return (int)command.ExecuteScalar();
                }
            }
        }
        // Tambah di class properties
        private List<(TimeSpan Start, TimeSpan End)> _cachedBreakTimes = null;

        // Ganti method GetAllBreakTimes
        private List<(TimeSpan Start, TimeSpan End)> GetAllBreakTimes()
        {
            // Kalau udah ada cache, return langsung
            if (_cachedBreakTimes != null)
                return _cachedBreakTimes;

            var breakTimes = new List<(TimeSpan Start, TimeSpan End)>();
            breakTimes.AddRange(FixedBreakTimes);

            if (!string.IsNullOrEmpty(AdditionalBreakTime1Start) && !string.IsNullOrEmpty(AdditionalBreakTime1End))
            {
                if (TryParseTimeSpan(AdditionalBreakTime1Start, out TimeSpan start1) &&
                    TryParseTimeSpan(AdditionalBreakTime1End, out TimeSpan end1))
                {
                    breakTimes.Add((start1, end1));
                }
            }

            if (!string.IsNullOrEmpty(AdditionalBreakTime2Start) && !string.IsNullOrEmpty(AdditionalBreakTime2End))
            {
                if (TryParseTimeSpan(AdditionalBreakTime2Start, out TimeSpan start2) &&
                    TryParseTimeSpan(AdditionalBreakTime2End, out TimeSpan end2))
                {
                    breakTimes.Add((start2, end2));
                }
            }

            // Simpan ke cache
            _cachedBreakTimes = breakTimes;
            return breakTimes;
        }
        private void EnsureValidCurrentPage() { if (TotalRecords == 0) { CurrentPage = 1; return; } int maxPages = (int)Math.Ceiling((double)TotalRecords / PageSize); if (CurrentPage > maxPages) CurrentPage = maxPages; else if (CurrentPage < 1) CurrentPage = 1; }

        private string CategorizeReason(string reason)
        {
            reason = reason?.ToLower() ?? "";
            if (reason.Contains("model changing loss")) return "Model Changing Loss";
            else if (reason.Contains("material shortage external")) return "Material Shortage External";
            else if (reason.Contains("man power adjustment")) return "Man Power Adjustment";
            else if (reason.Contains("material shortage internal")) return "Material Shortage Internal";
            else if (reason.Contains("material shortage inhouse")) return "Material Shortage Inhouse";
            else if (reason.Contains("quality trouble")) return "Quality Trouble";
            else if (reason.Contains("machine & tools trouble")) return "Machine & Tools Trouble";
            else if (reason.Contains("rework")) return "Rework";
            else if (reason.Contains("morning assembly")) return "Morning Assembly";
            else return "Reason Not Fill";
        }

        public int GetTotalDurationAllCategories() => CategorySummary.Values.Sum();
        public double SecondsToMinutes(int seconds) => Math.Round(seconds / 60.0, 2);
        public List<int> GetPageSizeOptions() => new List<int> { 10 };

        public IActionResult OnPostExportExcel()
        {
            LoadBreakTimeForToday();
            SetDatesFromMonthYear();
            var breakTimes = GetAllBreakTimes();

            // ✅ Pakai GetCombinedRecords yang sudah ada
            var exportData = GetCombinedRecords(
                StartSelectedDate, StartSelectedDate,
                StartSelectedDate, EndSelectedDate,
                breakTimes
            ).OrderByDescending(x => x.Date).ToList();

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Loss Time Data");
                worksheet.Cell(1, 1).Value = "No";
                worksheet.Cell(1, 2).Value = "Date";
                worksheet.Cell(1, 3).Value = "Category";
                worksheet.Cell(1, 4).Value = "Start Time";
                worksheet.Cell(1, 5).Value = "End Time";
                worksheet.Cell(1, 6).Value = "Duration (Sec)";
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
                    worksheet.Cell(row, 6).Value = item.Duration;
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
                        $"LossTime_{StartSelectedDate:yyyyMMdd}-{EndSelectedDate:yyyyMMdd}.xlsx");
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

        private string NormalizeCategoryName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "Uncategorized";

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            return textInfo.ToTitleCase(input.Trim().ToLower());
        }

        public async Task<IActionResult> OnPostImportExcelActualAsync()
        {

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                using (var stream = new MemoryStream())
                {
                    await UploadedExcel.CopyToAsync(stream);
                    using (var package = new ExcelPackage(stream))
                    {
                        var sheet = package.Workbook.Worksheets[0];
                        int rowCount = sheet.Dimension.Rows;

                        var newActuals = new List<LossTimeActual>();

                        for (int row = 3; row <= rowCount; row++)
                        {
                            var catName = sheet.Cells[row, 2].Value?.ToString()?.Trim();

                            if (string.IsNullOrEmpty(catName) ||
                                catName.ToLower().Contains("loss (min)") ||
                                catName.ToLower().Contains("loss category")) continue;

                            catName = NormalizeCategoryName(catName);

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

                        var dataToDelete = _context.LossTimeActuals
                            .Where(x => x.MachineLine == this.UploadMachineLine &&
                                        x.Month == TargetMonth &&
                                        x.Year == TargetYear);

                        _context.LossTimeActuals.RemoveRange(dataToDelete);

                        if (newActuals.Any())
                        {
                            _context.LossTimeActuals.AddRange(newActuals);
                            await _context.SaveChangesAsync();
                            TempData["Success"] = $"Berhasil import {newActuals.Count} data Actual untuk {UploadMachineLine} (Bulan {TargetMonth}).";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
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