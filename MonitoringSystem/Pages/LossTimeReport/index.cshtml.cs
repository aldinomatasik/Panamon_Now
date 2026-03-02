using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Globalization;
using OfficeOpenXml;
using MonitoringSystem.Models;
using MonitoringSystem.Data;
using System;
using System.Drawing;

namespace MonitoringSystem.Pages.LossTimeReport
{
    public class indexModel : PageModel
    {
        private readonly ApplicationDbContext _context;

        public indexModel(ApplicationDbContext context)
        {
            _context = context;
        }

        [BindProperty(SupportsGet = true)]
        public int SelectedYear { get; set; } = DateTime.Today.Year;

        [BindProperty(SupportsGet = true)]
        public string MachineLine { get; set; } = "All";

        [BindProperty]
        public string UploadMachineLine { get; set; }

        [BindProperty]
        public IFormFile UploadedExcel { get; set; }

        public string ChartDataJson { get; set; } = "{}";
        public List<string> Categories { get; set; } = new List<string>();
        public Dictionary<string, double[]> DetailActuals { get; set; } = new Dictionary<string, double[]>();
        public Dictionary<string, double[]> DetailPlans { get; set; } = new Dictionary<string, double[]>();
        public double[] TotalActualPerMonth { get; set; } = new double[12];
        public double[] TotalPlanPerMonth { get; set; } = new double[12];
        public double[] ActualRatios { get; set; } = new double[12];
        public double[] PlanRatios { get; set;} = new double[12];

        public void OnGet()
        {
            string[] months = { "April", "May", "June", "July", "August", "September", "October", "November", "December", "January", "February", "March" };

            var actualsRaw = GetDetailedActualData(SelectedYear, MachineLine);

            var planQuery = _context.LossTimePlans.AsQueryable();
            planQuery = planQuery.Where(x =>
                (x.Year == SelectedYear && x.Month >= 4) ||
                (x.Year == SelectedYear + 1 && x.Month <= 3)
            );

            if (MachineLine != "All")
            {
                planQuery = planQuery.Where(x => x.MachineLine == MachineLine);
            }

            var plansRaw = planQuery.ToList()
                .GroupBy(x => new {Category = NormalizeCategoryName(x.Category), Month = x.Month })
                .Select(g => new { Category = g.Key.Category, Month = g.Key.Month, Total = g.Sum(x => x.TargetMinutes) })
                .ToList();

            var plansRatioRaw = planQuery.ToList()
                .GroupBy(x => x.Month)
                .Select(g => new { Month = g.Key, RatioVal = g.Max(x => x.Ratio)  })
                .ToList();

            //Data not valid(need improvement)
            var workingTimeRaw = GetMonthlyWorkingTime(SelectedYear, MachineLine);

            var allCats = actualsRaw.Select(x => x.Category)
                          .Union(plansRaw.Select(x => x.Category))
                          .Distinct()
                          //.OrderBy(x => x)
                          .ToList();

            Categories = allCats
                .OrderBy(c => {
                    string group = GetCategoryGroup(c);

                    if (group == "Working Loss") return 1;
                    if (group == "Fixed Loss") return 2;
                    return 99;
                })
                .ThenBy(c => c)
                .ToList();

            foreach (var cat in Categories)
            {
                double[] actArr = new double[12];
                double[] planArr = new double[12];

                var catActuals = actualsRaw.Where(x => x.Category == cat);
                foreach (var item in catActuals)
                {
                    int arrayIndex = (item.Month - 4 + 12) % 12;
                    actArr[arrayIndex] = Math.Round(item.Total, 1);
                }

                var catPlans = plansRaw.Where(x => x.Category == cat);
                foreach (var item in catPlans)
                {
                    int arrayIndex = (item.Month - 4 + 12) % 12;
                    planArr[arrayIndex] = Math.Round(item.Total, 1);
                }

                DetailActuals.Add(cat, actArr);
                DetailPlans.Add(cat, planArr);
            }

            for (int i = 0; i < 12; i++)
            {
                TotalActualPerMonth[i] = DetailActuals.Sum(x => x.Value[i]);
                TotalPlanPerMonth[i] = DetailPlans.Sum(x => x.Value[i]);
            }

            for (int i = 0; i < 12; i++)
            {
                int monthNum = (i + 4) > 12 ? (i + 4) - 12 : (i + 4);
                var pRatio = plansRatioRaw.FirstOrDefault(x => x.Month == monthNum);
                PlanRatios[i] = pRatio != null ? (double)pRatio.RatioVal : 0;

                double totalLoss = TotalActualPerMonth[i];
                double workingTime = workingTimeRaw.ContainsKey(monthNum) ? workingTimeRaw[monthNum] : 0;

                if (workingTime > 0 )
                {
                    ActualRatios[i] = Math.Round((totalLoss / workingTime) * 100, 2);
                }
                else
                {
                    ActualRatios[i] = 0;
                }
            }

            var datasets = new List<object>();

            var chartPayload = new
            {
                Labels = months,
                Categories = Categories,
                Actuals = DetailActuals,
                Plans = DetailPlans,
                RatioActual = ActualRatios,
                RatioPlan = PlanRatios
            };

            ChartDataJson = System.Text.Json.JsonSerializer.Serialize(chartPayload);
        }

        public string GetCategoryGroup(string categoryName)
        {
            string lowerCat = categoryName.ToLower().Trim();

            if (lowerCat.Contains("break time (am/pm") || lowerCat.Contains("company activity") || lowerCat.Contains("morning assembly") || lowerCat.Contains("cleaning") || lowerCat.Contains("stock opname") || lowerCat.Contains("general assembly") || lowerCat.Contains("maintenance") || lowerCat.Contains("trial run") || lowerCat.Contains("training education") || lowerCat.Contains("free talking/qc activity") || lowerCat.Contains("no production day"))
                return "Fixed Loss";

            if (lowerCat.Contains("quality trouble") || lowerCat.Contains("model changing loss") || lowerCat.Contains("material shortage external") || lowerCat.Contains("machine & tools trouble") || lowerCat.Contains("man power adjustment") || lowerCat.Contains("material shortage inhouse") || lowerCat.Contains("material shortage internal") || lowerCat.Contains("set repairing loss") || lowerCat.Contains("gawse - external bodies") || lowerCat.Contains("rework") || lowerCat.Contains("mold changing loss"))
                return "Working Loss";

            return "Others";
        }

        private string NormalizeCategoryName(string input)
        {
            if (string.IsNullOrWhiteSpace(input)) return "Uncategorized";

            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            return textInfo.ToTitleCase(input.Trim().ToLower());
        }

        private class MonthlyCategoryData
        {
            public int Month { get; set; }
            public string Category { get; set; }
            public double Total { get; set; }
        }

        private List<MonthlyCategoryData> GetDetailedActualData(int fiscalYear, string line)
        {
            var result = new List<MonthlyCategoryData>();
            var connectionString = _context.Database.GetDbConnection().ConnectionString;

            DateTime startDate = new DateTime(fiscalYear, 4, 1);
            DateTime endDate = new DateTime(fiscalYear + 1, 3, 31);

            string query = @"
                SELECT 
                    MONTH(Date) as MonthVal,
                    Reason, 
                    SUM(LossTime) / 60 as TotalMinutes
                FROM AssemblyLossTime
                WHERE Date >= @Start AND Date <= @End
            ";

            if (line != "All")
            {
                query += " AND MachineCode = @MachineCode";
            }

            query += " GROUP BY MONTH(Date), Reason";

            try
            {
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Start", startDate);
                        cmd.Parameters.AddWithValue("@End", endDate);

                        if (line != "All")
                        {
                            cmd.Parameters.AddWithValue("@MachineCode", line);
                        }

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string rawCat = reader["Reason"]?.ToString();
                                string cleanCat = NormalizeCategoryName(rawCat);

                                result.Add(new MonthlyCategoryData
                                {
                                    Month = Convert.ToInt32(reader["MonthVal"]),
                                    Category = cleanCat,
                                    Total = reader["TotalMinutes"] != DBNull.Value ? Convert.ToDouble(reader["TotalMinutes"]) : 0
                                });
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("SQL Error: " + ex.Message);
            }

            return result;
        }

        //Not Valid Need Improvement for for calculation working time
        private Dictionary<int, double> GetMonthlyWorkingTime(int fiscalYear, string line)
        {
            var result = new Dictionary<int, double>();
            string query = @"
            SELECT 
                MONTH(Date) as MonthVal,
                SUM(WorkingTime) as TotalWT 
            FROM ProductionData 
            WHERE Date >= @Start AND Date <= @End
        ";

            if (line != "All") query += " AND MachineCode = @MachineCode";
            query += " GROUP BY MONTH(Date)";

            DateTime startDate = new DateTime(fiscalYear, 4, 1);
            DateTime endDate = new DateTime(fiscalYear + 1, 3, 31);

            try
            {
                using (var conn = new SqlConnection(_context.Database.GetDbConnection().ConnectionString))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@Start", startDate);
                        cmd.Parameters.AddWithValue("@End", endDate);
                        if (line != "All") cmd.Parameters.AddWithValue("@MachineCode", line);

                        using (var reader = cmd.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                int m = Convert.ToInt32(reader["MonthVal"]);
                                double val = reader["TotalWT"] != DBNull.Value ? Convert.ToDouble(reader["TotalWT"]) : 0;
                                if (!result.ContainsKey(m)) result.Add(m, val);
                            }
                        }
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("WT Error: " + ex.Message); }
            return result;
        }

        public async Task<IActionResult> OnPostImportExcelAsync()
        {
            if (UploadedExcel == null || UploadedExcel.Length == 0)
            {
                TempData["Error"] = "File Excel belum dipilih.";
                return RedirectToPage(new { SelectedYear, MachineLine });
            }

            if (UploadMachineLine == "All" || string.IsNullOrEmpty(UploadMachineLine))
            {
                TempData["Error"] = "Saat Upload, anda harus memilih Mesin Spesifik (CU atau CS) di dalam Pop-up.";
                return RedirectToPage(new { SelectedYear, MachineLine });
            }

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

                        var monthColMap = new Dictionary<int, (int targetMinutes, int ratioPlans)>
                        {
                            { 4, (3,4) }, { 5, (5, 6) }, { 6, (7, 8) }, { 7, (9, 10) }, { 8, (11, 12) }, { 9, (13, 14) },
                            { 10, (15, 16) }, { 11, (17, 18) }, { 12, (19, 20) }, { 1, (21, 22) }, { 2, (23, 24) }, {3, (25, 26) }
                         };

                        var newPlans = new List<LossTimePlan>();

                        for (int row = 4; row <= rowCount; row++)
                        {
                            var catName = sheet.Cells[row, 2].Text?.Trim();
                            if (string.IsNullOrEmpty(catName) || catName.ToLower().Contains("loss category")) continue;

                            catName = NormalizeCategoryName(catName);

                            foreach (var map in monthColMap)
                            {
                                var targetText = sheet.Cells[row, map.Value.targetMinutes].Text;
                                var ratioCell = sheet.Cells[row, map.Value.ratioPlans].Value;

                                bool targetValid = double.TryParse(targetText, out double targetVal) && targetVal >= 0;

                                decimal ratioVal = 0;
                                bool ratioValid = false;

                                if (ratioCell != null)
                                {
                                    ratioVal = Convert.ToDecimal(ratioCell) * 100;
                                    ratioValid = ratioVal >= 0;
                                }

                                if (!targetValid && !ratioValid) continue;

                                {
                                    int dataYear = (map.Key >= 4) ? SelectedYear : SelectedYear + 1;

                                    newPlans.Add(new LossTimePlan
                                    {
                                        Category = catName,
                                        MachineLine = this.UploadMachineLine,
                                        Month = map.Key,
                                        Year = dataYear,
                                        TargetMinutes = targetValid ? targetVal : 0,
                                        Ratio = ratioValid ? ratioVal : 0
                                    });
                                }
                            }
                        }

                        var dataToDelete = _context.LossTimePlans
                            .Where(x => x.MachineLine == this.UploadMachineLine &&
                                        ((x.Year == SelectedYear && x.Month >= 4) ||
                                         (x.Year == SelectedYear + 1 && x.Month <= 3)));

                        _context.LossTimePlans.RemoveRange(dataToDelete);

                        if (newPlans.Any())
                        {
                            _context.LossTimePlans.AddRange(newPlans);
                            await _context.SaveChangesAsync();
                            TempData["Success"] = $"Berhasil import Plan untuk {UploadMachineLine}.";
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TempData["Error"] = "Gagal Import: " + ex.Message;
            }

            return RedirectToPage(new { SelectedYear, MachineLine });
        }

        public IActionResult OnGetDownloadTemplate()
        {
            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "data", "planlosstime", "LossTimePlan_Template.xlsx");
            if (!System.IO.File.Exists(filePath)) return NotFound("File template tidak ditemukan di server.");
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "LossTimePlan_Template.xlsx");
        }
    }
}
