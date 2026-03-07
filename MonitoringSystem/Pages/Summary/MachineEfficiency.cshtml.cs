using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using OfficeOpenXml;

namespace MonitoringSystem.Pages.Summary
{
    public class MachineEfficiencyModel : PageModel
    {
        private readonly IConfiguration _config;
        private readonly IWebHostEnvironment _env;

        public MachineEfficiencyModel(IConfiguration config, IWebHostEnvironment env)
        {
            _config = config;
            _env = env;
        }

        public void OnGet() { }

        // =====================================================
        // API: /Summary/MachineEfficiency?handler=Machines
        // =====================================================
        public JsonResult OnGetMachines()
        {
            var machines = new List<object>();
            try
            {
                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"SELECT ID, MachineName FROM [dbo].[Machine] ORDER BY MachineName";
                using var cmd = new SqlCommand(sql, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    machines.Add(new { value = reader["ID"].ToString(), label = reader["MachineName"].ToString() });
            }
            catch (Exception ex) { return new JsonResult(new { error = ex.Message }); }
            return new JsonResult(machines);
        }

        // =====================================================
        // API: ?handler=WorkingTime&machineName=X&date=Y&shift=Z
        // =====================================================
        public JsonResult OnGetWorkingTime(string machineName, DateTime date, string shift)
        {
            try
            {
                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"
                    SELECT TOP 1 WorkingTime
                    FROM [dbo].[MachineEfficiency]
                    WHERE MachineName = @MachineName
                      AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                      AND Shift = @Shift
                      AND WorkingTime IS NOT NULL
                    ORDER BY ID DESC";
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@MachineName", machineName ?? "");
                cmd.Parameters.AddWithValue("@Date", date);
                cmd.Parameters.AddWithValue("@Shift", shift ?? "");
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                    return new JsonResult(new { found = true, workingTime = Convert.ToDouble(result) });
                return new JsonResult(new { found = false });
            }
            catch (Exception ex) { return new JsonResult(new { found = false, error = ex.Message }); }
        }

        // =====================================================
        // API: ?handler=LoadExisting&machineName=X&date=Y&shift=Z
        // Fetch semua data existing untuk pre-fill form
        // Toleran terhadap berbagai format Shift di DB:
        //   "Shift 1" / "1" / NULL (semua dianggap Shift 1), dst.
        // =====================================================
        public JsonResult OnGetLoadExisting(string machineName, DateTime date, string shift)
        {
            try
            {
                // Normalisasi shift: toleran format lama ("1","2","3","NS") & baru ("Shift 1" dst) & NULL
                string sNew = shift switch
                {
                    "1" or "Shift 1" => "Shift 1",
                    "2" or "Shift 2" => "Shift 2",
                    "3" or "Shift 3" => "Shift 3",
                    "NS" or "Non Shift" => "Non Shift",
                    var s => s
                };
                string sOld = sNew switch
                {
                    "Shift 1" => "1",
                    "Shift 2" => "2",
                    "Shift 3" => "3",
                    "Non Shift" => "NS",
                    var s => s
                };

                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"
                    SELECT TOP 1
                        PlanQty, DefectQty, GoodProductionQty, WorkingTime,
                        QualityTrouble, ModelChangingLoss, MaterialShortageExternal,
                        MachineToolsTrouble, ManPowerAdjustment, MaterialShortageInhouse,
                        MaterialShortageInternal, SetRepairingLoss, GawseExternalBodies,
                        Rework, MoldChangingLoss, BreakTime, CompanyActivity, MorningAssembly,
                        Cleaning, StockOpname, GeneralAssembly, Maintenance, TrialRun,
                        TrainingEducation, FreeTalkingQC, NoProductionDay
                    FROM [dbo].[MachineEfficiency]
                    WHERE MachineName = @MachineName
                      AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                      AND (Shift = @ShiftNew OR Shift = @ShiftOld OR Shift IS NULL)
                    ORDER BY ID DESC";
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@MachineName", machineName ?? "");
                cmd.Parameters.AddWithValue("@Date", date);
                cmd.Parameters.AddWithValue("@ShiftNew", sNew);
                cmd.Parameters.AddWithValue("@ShiftOld", sOld);
                using var reader = cmd.ExecuteReader();
                if (reader.Read())
                {
                    double? G(string col) => reader[col] == DBNull.Value ? null : Convert.ToDouble(reader[col]);
                    return new JsonResult(new
                    {
                        found = true,
                        planQty = G("PlanQty"),
                        defectQty = G("DefectQty"),
                        goodProductionQty = G("GoodProductionQty"),
                        workingTime = G("WorkingTime"),
                        qualityTrouble = G("QualityTrouble"),
                        modelChangingLoss = G("ModelChangingLoss"),
                        materialShortageExternal = G("MaterialShortageExternal"),
                        machineToolsTrouble = G("MachineToolsTrouble"),
                        manPowerAdjustment = G("ManPowerAdjustment"),
                        materialShortageInhouse = G("MaterialShortageInhouse"),
                        materialShortageInternal = G("MaterialShortageInternal"),
                        setRepairingLoss = G("SetRepairingLoss"),
                        gawseExternalBodies = G("GawseExternalBodies"),
                        rework = G("Rework"),
                        moldChangingLoss = G("MoldChangingLoss"),
                        breakTime = G("BreakTime"),
                        companyActivity = G("CompanyActivity"),
                        morningAssembly = G("MorningAssembly"),
                        cleaning = G("Cleaning"),
                        stockOpname = G("StockOpname"),
                        generalAssembly = G("GeneralAssembly"),
                        maintenance = G("Maintenance"),
                        trialRun = G("TrialRun"),
                        trainingEducation = G("TrainingEducation"),
                        freeTalkingQC = G("FreeTalkingQC"),
                        noProductionDay = G("NoProductionDay")
                    });
                }
                return new JsonResult(new { found = false });
            }
            catch (Exception ex) { return new JsonResult(new { found = false, error = ex.Message }); }
        }

        // =====================================================
        // API: /Summary/MachineEfficiency?handler=LossCategories
        // =====================================================
        public JsonResult OnGetLossCategories()
        {
            var workingLossCols = new List<string>
            {
                "QualityTrouble", "ModelChangingLoss", "MaterialShortageExternal",
                "MachineToolsTrouble", "ManPowerAdjustment", "MaterialShortageInhouse",
                "MaterialShortageInternal", "SetRepairingLoss", "GawseExternalBodies",
                "Rework", "MoldChangingLoss"
            };
            var fixedLossCols = new List<string>
            {
                "BreakTime", "CompanyActivity", "MorningAssembly", "Cleaning",
                "StockOpname", "GeneralAssembly", "Maintenance", "TrialRun",
                "TrainingEducation", "FreeTalkingQC", "NoProductionDay"
            };
            var existingCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            try
            {
                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA='dbo' AND TABLE_NAME='MachineEfficiency'";
                using var cmd = new SqlCommand(sql, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) existingCols.Add(reader.GetString(0));
            }
            catch
            {
                existingCols.UnionWith(workingLossCols);
                existingCols.UnionWith(fixedLossCols);
            }
            static string ToLabel(string col) =>
                System.Text.RegularExpressions.Regex.Replace(col, @"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ");
            var workingLoss = workingLossCols.Where(c => existingCols.Contains(c)).Select(c => new { value = c, label = ToLabel(c) }).ToArray();
            var fixedLoss = fixedLossCols.Where(c => existingCols.Contains(c)).Select(c => new { value = c, label = ToLabel(c) }).ToArray();
            return new JsonResult(new { workingLoss, fixedLoss });
        }

        // =====================================================
        // HELPER: tambah parameter ke SqlCommand
        // =====================================================
        private static void AddAllParams(SqlCommand cmd, MachineEfficiencyInput input,
            double? quality, double? operatingRatio, double? ability, double? oee,
            double? achievement, string? category)
        {
            cmd.Parameters.AddWithValue("@MachineName", input.MachineName ?? "");
            cmd.Parameters.AddWithValue("@Date", input.Date);
            cmd.Parameters.AddWithValue("@Shift", input.Shift ?? "");
            cmd.Parameters.AddWithValue("@PlanQty", (object?)input.PlanQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DefectQty", (object?)input.DefectQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@GoodProductionQty", (object?)input.GoodProductionQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@WorkingTime", (object?)input.WorkingTime ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Quality", (object?)quality ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OperatingRatio", (object?)operatingRatio ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ability", (object?)ability ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OEE", (object?)oee ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Achievement", (object?)achievement ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Category", (object?)category ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@QualityTrouble", (object?)input.QualityTrouble ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ModelChangingLoss", (object?)input.ModelChangingLoss ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MaterialShortageExternal", (object?)input.MaterialShortageExternal ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MachineToolsTrouble", (object?)input.MachineToolsTrouble ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@ManPowerAdjustment", (object?)input.ManPowerAdjustment ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MaterialShortageInhouse", (object?)input.MaterialShortageInhouse ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MaterialShortageInternal", (object?)input.MaterialShortageInternal ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@SetRepairingLoss", (object?)input.SetRepairingLoss ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@GawseExternalBodies", (object?)input.GawseExternalBodies ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Rework", (object?)input.Rework ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MoldChangingLoss", (object?)input.MoldChangingLoss ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@BreakTime", (object?)input.BreakTime ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@CompanyActivity", (object?)input.CompanyActivity ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@MorningAssembly", (object?)input.MorningAssembly ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Cleaning", (object?)input.Cleaning ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@StockOpname", (object?)input.StockOpname ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@GeneralAssembly", (object?)input.GeneralAssembly ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Maintenance", (object?)input.Maintenance ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TrialRun", (object?)input.TrialRun ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@TrainingEducation", (object?)input.TrainingEducation ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@FreeTalkingQC", (object?)input.FreeTalkingQC ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@NoProductionDay", (object?)input.NoProductionDay ?? DBNull.Value);
        }

        // =====================================================
        // HELPER: hitung semua metric dari input
        // =====================================================
        private static (double? quality, double? operatingRatio, double? ability, double? oee, double? achievement, string? category)
            CalcMetrics(MachineEfficiencyInput input)
        {
            double? quality = null;
            if (input.GoodProductionQty.HasValue && input.GoodProductionQty.Value > 0)
            {
                double defect = input.DefectQty ?? 0;
                quality = Math.Round(((input.GoodProductionQty.Value - defect) / input.GoodProductionQty.Value) * 100.0, 2);
            }

            double? operatingRatio = null;
            if (input.WorkingTime.HasValue && input.WorkingTime.Value > 0)
            {
                double totalLoss =
                    (input.QualityTrouble ?? 0) + (input.ModelChangingLoss ?? 0) +
                    (input.MaterialShortageExternal ?? 0) + (input.MachineToolsTrouble ?? 0) +
                    (input.ManPowerAdjustment ?? 0) + (input.MaterialShortageInhouse ?? 0) +
                    (input.MaterialShortageInternal ?? 0) + (input.SetRepairingLoss ?? 0) +
                    (input.GawseExternalBodies ?? 0) + (input.Rework ?? 0) +
                    (input.MoldChangingLoss ?? 0) + (input.BreakTime ?? 0) +
                    (input.CompanyActivity ?? 0) + (input.MorningAssembly ?? 0) +
                    (input.Cleaning ?? 0) + (input.StockOpname ?? 0) +
                    (input.GeneralAssembly ?? 0) + (input.Maintenance ?? 0) +
                    (input.TrialRun ?? 0) + (input.TrainingEducation ?? 0) +
                    (input.FreeTalkingQC ?? 0) + (input.NoProductionDay ?? 0);
                operatingRatio = Math.Round(((input.WorkingTime.Value - totalLoss) / input.WorkingTime.Value) * 100.0, 2);
            }

            double? ability = null;
            if (input.PlanQty.HasValue && input.PlanQty.Value > 0 && input.GoodProductionQty.HasValue)
                ability = Math.Round((input.GoodProductionQty.Value / input.PlanQty.Value) * 100.0, 2);

            double? oee = null;
            if (operatingRatio.HasValue && ability.HasValue && quality.HasValue)
                oee = Math.Round((operatingRatio.Value * ability.Value * quality.Value) / 10000.0, 2);

            double? achievement = null;
            if (input.PlanQty.HasValue && input.PlanQty.Value > 0 && input.GoodProductionQty.HasValue)
                achievement = Math.Round((input.GoodProductionQty.Value / input.PlanQty.Value) * 100.0, 2);

            string? category = oee.HasValue
                ? (oee.Value >= 85 ? "Good" : oee.Value >= 60 ? "Average" : "Poor")
                : null;

            return (quality, operatingRatio, ability, oee, achievement, category);
        }

        // =====================================================
        // API: /Summary/MachineEfficiency?handler=Submit
        // UPSERT: UPDATE jika MachineName+Date+Shift sudah ada, INSERT jika belum
        // =====================================================
        public JsonResult OnPostSubmit([FromBody] MachineEfficiencyInput input)
        {
            try
            {
                var (quality, operatingRatio, ability, oee, achievement, category) = CalcMetrics(input);

                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();

                // Normalisasi shift: frontend baru kirim "Shift 1" dst,
                // tapi data lama di DB mungkin "1"/"2"/"3"/"NS"/NULL.
                // Selalu simpan dalam format baru, tapi saat query cari semua alias.
                string shiftNorm = (input.Shift ?? "") switch
                {
                    "1" or "Shift 1" => "Shift 1",
                    "2" or "Shift 2" => "Shift 2",
                    "3" or "Shift 3" => "Shift 3",
                    "NS" or "Non Shift" => "Non Shift",
                    var s => s
                };
                input.Shift = shiftNorm; // pastikan INSERT/UPDATE pakai format baru

                string shiftOld = shiftNorm switch
                {
                    "Shift 1" => "1",
                    "Shift 2" => "2",
                    "Shift 3" => "3",
                    "Non Shift" => "NS",
                    var s => s
                };

                int existingCount = 0;
                using (var checkCmd = new SqlCommand(@"
                    SELECT COUNT(1) FROM [dbo].[MachineEfficiency]
                    WHERE [MachineName] = @MachineName
                      AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                      AND ([Shift] = @ShiftNew OR [Shift] = @ShiftOld OR [Shift] IS NULL)", conn))
                {
                    checkCmd.Parameters.AddWithValue("@MachineName", input.MachineName ?? "");
                    checkCmd.Parameters.AddWithValue("@Date", input.Date);
                    checkCmd.Parameters.AddWithValue("@ShiftNew", shiftNorm);
                    checkCmd.Parameters.AddWithValue("@ShiftOld", shiftOld);
                    existingCount = (int)checkCmd.ExecuteScalar();
                }

                string sql;
                if (existingCount > 0)
                {
                    sql = @"
                        UPDATE [dbo].[MachineEfficiency] SET
                            [PlanQty]                  = @PlanQty,
                            [DefectQty]                = @DefectQty,
                            [GoodProductionQty]        = @GoodProductionQty,
                            [WorkingTime]              = @WorkingTime,
                            [Quality]                  = @Quality,
                            [OperatingRatio]           = @OperatingRatio,
                            [Ability]                  = @Ability,
                            [OEE]                      = @OEE,
                            [Achievement]              = @Achievement,
                            [Category]                 = @Category,
                            [QualityTrouble]           = @QualityTrouble,
                            [ModelChangingLoss]        = @ModelChangingLoss,
                            [MaterialShortageExternal] = @MaterialShortageExternal,
                            [MachineToolsTrouble]      = @MachineToolsTrouble,
                            [ManPowerAdjustment]       = @ManPowerAdjustment,
                            [MaterialShortageInhouse]  = @MaterialShortageInhouse,
                            [MaterialShortageInternal] = @MaterialShortageInternal,
                            [SetRepairingLoss]         = @SetRepairingLoss,
                            [GawseExternalBodies]      = @GawseExternalBodies,
                            [Rework]                   = @Rework,
                            [MoldChangingLoss]         = @MoldChangingLoss,
                            [BreakTime]                = @BreakTime,
                            [CompanyActivity]          = @CompanyActivity,
                            [MorningAssembly]          = @MorningAssembly,
                            [Cleaning]                 = @Cleaning,
                            [StockOpname]              = @StockOpname,
                            [GeneralAssembly]          = @GeneralAssembly,
                            [Maintenance]              = @Maintenance,
                            [TrialRun]                 = @TrialRun,
                            [TrainingEducation]        = @TrainingEducation,
                            [FreeTalkingQC]            = @FreeTalkingQC,
                            [NoProductionDay]          = @NoProductionDay
                        WHERE [MachineName] = @MachineName
                          AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                          AND ([Shift] = @ShiftNew OR [Shift] = @ShiftOld OR [Shift] IS NULL)";
                }
                else
                {
                    sql = @"
                        INSERT INTO [dbo].[MachineEfficiency]
                            ([MachineName],[Date],[Shift],
                             [PlanQty],[DefectQty],[GoodProductionQty],[WorkingTime],
                             [Quality],[OperatingRatio],[Ability],[OEE],[Achievement],[Category],
                             [QualityTrouble],[ModelChangingLoss],[MaterialShortageExternal],
                             [MachineToolsTrouble],[ManPowerAdjustment],[MaterialShortageInhouse],
                             [MaterialShortageInternal],[SetRepairingLoss],[GawseExternalBodies],
                             [Rework],[MoldChangingLoss],
                             [BreakTime],[CompanyActivity],[MorningAssembly],[Cleaning],
                             [StockOpname],[GeneralAssembly],[Maintenance],[TrialRun],
                             [TrainingEducation],[FreeTalkingQC],[NoProductionDay])
                        VALUES
                            (@MachineName,@Date,@Shift,
                             @PlanQty,@DefectQty,@GoodProductionQty,@WorkingTime,
                             @Quality,@OperatingRatio,@Ability,@OEE,@Achievement,@Category,
                             @QualityTrouble,@ModelChangingLoss,@MaterialShortageExternal,
                             @MachineToolsTrouble,@ManPowerAdjustment,@MaterialShortageInhouse,
                             @MaterialShortageInternal,@SetRepairingLoss,@GawseExternalBodies,
                             @Rework,@MoldChangingLoss,
                             @BreakTime,@CompanyActivity,@MorningAssembly,@Cleaning,
                             @StockOpname,@GeneralAssembly,@Maintenance,@TrialRun,
                             @TrainingEducation,@FreeTalkingQC,@NoProductionDay)";
                }

                using var cmd = new SqlCommand(sql, conn);
                AddAllParams(cmd, input, quality, operatingRatio, ability, oee, achievement, category);
                // Untuk WHERE clause UPDATE: tambahkan ShiftNew & ShiftOld
                if (existingCount > 0)
                {
                    cmd.Parameters.AddWithValue("@ShiftNew", shiftNorm);
                    cmd.Parameters.AddWithValue("@ShiftOld", shiftOld);
                }
                cmd.ExecuteNonQuery();

                return new JsonResult(new { success = true, action = existingCount > 0 ? "updated" : "inserted" });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, error = ex.Message });
            }
        }

        // =====================================================
        // API: ?handler=DownloadTemplate
        // =====================================================
        public IActionResult OnGetDownloadTemplate()
        {
            var possibleNames = new[] { "TemplateMachine.xlsx", "Machine_input_template.xlsx", "template.xlsx" };
            string? filePath = null;
            foreach (var name in possibleNames)
            {
                var candidate = Path.Combine(_env.WebRootPath, "data", "MachineEfficiency", name);
                if (System.IO.File.Exists(candidate)) { filePath = candidate; break; }
            }
            if (filePath == null)
                return new NotFoundObjectResult(new { error = "File template tidak ditemukan di server." });
            var bytes = System.IO.File.ReadAllBytes(filePath);
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TemplateMachine.xlsx");
        }

        // =====================================================
        // API: ?handler=ImportExcel  (POST, multipart/form-data)
        // =====================================================
        public async Task<JsonResult> OnPostImportExcel(IFormFile file, string machineName, int month, int year)
        {
            if (file == null || file.Length == 0)
                return new JsonResult(new { success = false, error = "File tidak boleh kosong." });
            if (string.IsNullOrEmpty(machineName))
                return new JsonResult(new { success = false, error = "Nama machine harus dipilih." });

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var shifts = new[] { "Shift 1", "Shift 2", "Shift 3", "Non Shift" };

            const int ROW_QUALITY_TROUBLE = 3;
            const int ROW_MODEL_CHANGING_LOSS = 4;
            const int ROW_MATERIAL_SHORTAGE_EXT = 5;
            const int ROW_MACHINE_TOOLS_TROUBLE = 6;
            const int ROW_MAN_POWER_ADJUSTMENT = 7;
            const int ROW_MATERIAL_SHORTAGE_INHOUSE = 8;
            const int ROW_MATERIAL_SHORTAGE_INTERNAL = 9;
            const int ROW_SET_REPAIRING_LOSS = 10;
            const int ROW_GAWSE_EXTERNAL_BODIES = 11;
            const int ROW_REWORK = 12;
            const int ROW_MOLD_CHANGING_LOSS = 13;
            const int ROW_BREAK_TIME = 15;
            const int ROW_COMPANY_ACTIVITY = 16;
            const int ROW_MORNING_ASSEMBLY = 17;
            const int ROW_CLEANING = 18;
            const int ROW_STOCK_OPNAME = 19;
            const int ROW_GENERAL_ASSEMBLY = 20;
            const int ROW_MAINTENANCE = 21;
            const int ROW_TRIAL_RUN = 22;
            const int ROW_TRAINING_EDUCATION = 23;
            const int ROW_FREE_TALKING_QC = 24;
            const int ROW_NO_PRODUCTION_DAY = 25;
            const int ROW_PLAN = 27;
            const int ROW_GOOD_PRODUCTION_QTY = 28;
            const int ROW_DEFECT_QTY = 29;
            const int ROW_WORKING_TIME = 30;

            var newData = new List<MachineEfficiencyInput>();

            try
            {
                using var stream = new MemoryStream();
                await file.CopyToAsync(stream);
                using var package = new ExcelPackage(stream);
                var sheet = package.Workbook.Worksheets[0];
                int daysInMonth = DateTime.DaysInMonth(year, month);

                for (int day = 1; day <= daysInMonth; day++)
                {
                    for (int shiftIndex = 0; shiftIndex < 4; shiftIndex++)
                    {
                        int col = 3 + (day - 1) * 4 + shiftIndex;

                        double? planQty = TryParseDouble(sheet.Cells[ROW_PLAN, col].Value);
                        double? goodProductionQty = TryParseDouble(sheet.Cells[ROW_GOOD_PRODUCTION_QTY, col].Value);
                        double? defectQty = TryParseDouble(sheet.Cells[ROW_DEFECT_QTY, col].Value);
                        double? workingTime = TryParseDouble(sheet.Cells[ROW_WORKING_TIME, col].Value);
                        double? qualityTrouble = TryParseDouble(sheet.Cells[ROW_QUALITY_TROUBLE, col].Value);
                        double? modelChangingLoss = TryParseDouble(sheet.Cells[ROW_MODEL_CHANGING_LOSS, col].Value);
                        double? materialShortageExt = TryParseDouble(sheet.Cells[ROW_MATERIAL_SHORTAGE_EXT, col].Value);
                        double? machineToolsTrouble = TryParseDouble(sheet.Cells[ROW_MACHINE_TOOLS_TROUBLE, col].Value);
                        double? manPowerAdjustment = TryParseDouble(sheet.Cells[ROW_MAN_POWER_ADJUSTMENT, col].Value);
                        double? materialShortageInhouse = TryParseDouble(sheet.Cells[ROW_MATERIAL_SHORTAGE_INHOUSE, col].Value);
                        double? materialShortageInternal = TryParseDouble(sheet.Cells[ROW_MATERIAL_SHORTAGE_INTERNAL, col].Value);
                        double? setRepairingLoss = TryParseDouble(sheet.Cells[ROW_SET_REPAIRING_LOSS, col].Value);
                        double? gawseExternalBodies = TryParseDouble(sheet.Cells[ROW_GAWSE_EXTERNAL_BODIES, col].Value);
                        double? rework = TryParseDouble(sheet.Cells[ROW_REWORK, col].Value);
                        double? moldChangingLoss = TryParseDouble(sheet.Cells[ROW_MOLD_CHANGING_LOSS, col].Value);
                        double? breakTime = TryParseDouble(sheet.Cells[ROW_BREAK_TIME, col].Value);
                        double? companyActivity = TryParseDouble(sheet.Cells[ROW_COMPANY_ACTIVITY, col].Value);
                        double? morningAssembly = TryParseDouble(sheet.Cells[ROW_MORNING_ASSEMBLY, col].Value);
                        double? cleaning = TryParseDouble(sheet.Cells[ROW_CLEANING, col].Value);
                        double? stockOpname = TryParseDouble(sheet.Cells[ROW_STOCK_OPNAME, col].Value);
                        double? generalAssembly = TryParseDouble(sheet.Cells[ROW_GENERAL_ASSEMBLY, col].Value);
                        double? maintenance = TryParseDouble(sheet.Cells[ROW_MAINTENANCE, col].Value);
                        double? trialRun = TryParseDouble(sheet.Cells[ROW_TRIAL_RUN, col].Value);
                        double? trainingEducation = TryParseDouble(sheet.Cells[ROW_TRAINING_EDUCATION, col].Value);
                        double? freeTalkingQC = TryParseDouble(sheet.Cells[ROW_FREE_TALKING_QC, col].Value);
                        double? noProductionDay = TryParseDouble(sheet.Cells[ROW_NO_PRODUCTION_DAY, col].Value);

                        bool hasAnyData =
                            planQty.HasValue || goodProductionQty.HasValue || defectQty.HasValue || workingTime.HasValue ||
                            qualityTrouble.HasValue || modelChangingLoss.HasValue || materialShortageExt.HasValue ||
                            machineToolsTrouble.HasValue || manPowerAdjustment.HasValue || materialShortageInhouse.HasValue ||
                            materialShortageInternal.HasValue || setRepairingLoss.HasValue || gawseExternalBodies.HasValue ||
                            rework.HasValue || moldChangingLoss.HasValue || breakTime.HasValue || companyActivity.HasValue ||
                            morningAssembly.HasValue || cleaning.HasValue || stockOpname.HasValue || generalAssembly.HasValue ||
                            maintenance.HasValue || trialRun.HasValue || trainingEducation.HasValue ||
                            freeTalkingQC.HasValue || noProductionDay.HasValue;

                        if (!hasAnyData) continue;

                        newData.Add(new MachineEfficiencyInput
                        {
                            MachineName = machineName,
                            Date = new DateTime(year, month, day),
                            Shift = shifts[shiftIndex],
                            PlanQty = planQty,
                            GoodProductionQty = goodProductionQty,
                            DefectQty = defectQty,
                            WorkingTime = workingTime,
                            QualityTrouble = qualityTrouble,
                            ModelChangingLoss = modelChangingLoss,
                            MaterialShortageExternal = materialShortageExt,
                            MachineToolsTrouble = machineToolsTrouble,
                            ManPowerAdjustment = manPowerAdjustment,
                            MaterialShortageInhouse = materialShortageInhouse,
                            MaterialShortageInternal = materialShortageInternal,
                            SetRepairingLoss = setRepairingLoss,
                            GawseExternalBodies = gawseExternalBodies,
                            Rework = rework,
                            MoldChangingLoss = moldChangingLoss,
                            BreakTime = breakTime,
                            CompanyActivity = companyActivity,
                            MorningAssembly = morningAssembly,
                            Cleaning = cleaning,
                            StockOpname = stockOpname,
                            GeneralAssembly = generalAssembly,
                            Maintenance = maintenance,
                            TrialRun = trialRun,
                            TrainingEducation = trainingEducation,
                            FreeTalkingQC = freeTalkingQC,
                            NoProductionDay = noProductionDay,
                        });
                    }
                }

                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();

                using (var deleteCmd = new SqlCommand(@"
                    DELETE FROM [dbo].[MachineEfficiency]
                    WHERE [MachineName]=@MachineName AND MONTH([Date])=@Month AND YEAR([Date])=@Year", conn))
                {
                    deleteCmd.Parameters.AddWithValue("@MachineName", machineName);
                    deleteCmd.Parameters.AddWithValue("@Month", month);
                    deleteCmd.Parameters.AddWithValue("@Year", year);
                    deleteCmd.ExecuteNonQuery();
                }

                int insertedCount = 0;
                foreach (var item in newData)
                {
                    var (quality, operatingRatio, ability, oee, achievement, category) = CalcMetrics(item);

                    using var insertCmd = new SqlCommand(@"
                        INSERT INTO [dbo].[MachineEfficiency]
                            ([MachineName],[Date],[Shift],
                             [PlanQty],[DefectQty],[GoodProductionQty],[WorkingTime],
                             [Quality],[OperatingRatio],[Ability],[OEE],[Achievement],[Category],
                             [QualityTrouble],[ModelChangingLoss],[MaterialShortageExternal],
                             [MachineToolsTrouble],[ManPowerAdjustment],[MaterialShortageInhouse],
                             [MaterialShortageInternal],[SetRepairingLoss],[GawseExternalBodies],
                             [Rework],[MoldChangingLoss],
                             [BreakTime],[CompanyActivity],[MorningAssembly],[Cleaning],
                             [StockOpname],[GeneralAssembly],[Maintenance],[TrialRun],
                             [TrainingEducation],[FreeTalkingQC],[NoProductionDay])
                        VALUES
                            (@MachineName,@Date,@Shift,
                             @PlanQty,@DefectQty,@GoodProductionQty,@WorkingTime,
                             @Quality,@OperatingRatio,@Ability,@OEE,@Achievement,@Category,
                             @QualityTrouble,@ModelChangingLoss,@MaterialShortageExternal,
                             @MachineToolsTrouble,@ManPowerAdjustment,@MaterialShortageInhouse,
                             @MaterialShortageInternal,@SetRepairingLoss,@GawseExternalBodies,
                             @Rework,@MoldChangingLoss,
                             @BreakTime,@CompanyActivity,@MorningAssembly,@Cleaning,
                             @StockOpname,@GeneralAssembly,@Maintenance,@TrialRun,
                             @TrainingEducation,@FreeTalkingQC,@NoProductionDay)", conn);

                    AddAllParams(insertCmd, item, quality, operatingRatio, ability, oee, achievement, category);
                    insertCmd.ExecuteNonQuery();
                    insertedCount++;
                }

                return new JsonResult(new { success = true, message = $"Berhasil import {insertedCount} data untuk {machineName} (Bulan {month}/{year})." });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, error = "Gagal Import: " + ex.Message });
            }
        }

        private static double? TryParseDouble(object? val)
        {
            if (val == null) return null;
            return double.TryParse(val.ToString(), out double result) ? result : null;
        }

        // =====================================================
        // API: ?handler=MachineList
        // =====================================================
        public JsonResult OnGetMachineList()
        {
            var result = new List<object>();
            try
            {
                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"SELECT DISTINCT [MachineName] FROM [dbo].[MachineEfficiency] WHERE [MachineName] IS NOT NULL AND [MachineName]!='' ORDER BY [MachineName]";
                using var cmd = new SqlCommand(sql, conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read()) result.Add(new { machineName = reader["MachineName"]?.ToString() ?? "-" });
                return new JsonResult(result);
            }
            catch (Exception ex) { return new JsonResult(new { error = ex.Message }); }
        }

        // =====================================================
        // API: ?handler=Efficiency&month=X&year=Y
        // =====================================================
        public JsonResult OnGetEfficiency(int month, int year)
        {
            if (month < 1 || month > 12) return new JsonResult(new { error = "Bulan tidak valid (1-12)." });
            if (year < 2000 || year > 2100) return new JsonResult(new { error = "Tahun tidak valid." });

            var result = new List<object>();
            try
            {
                string connStr = _config.GetConnectionString("DefaultConnection")!;
                using var conn = new SqlConnection(connStr);
                conn.Open();
                string sql = @"
                    SELECT [MachineName],
                        ISNULL(CAST([OEE] AS VARCHAR),'-') AS [OEE],
                        ISNULL(CAST([OperatingRatio] AS VARCHAR),'-') AS [OperatingRatio],
                        ISNULL(CAST([Ability] AS VARCHAR),'-') AS [Ability],
                        ISNULL(CAST([Quality] AS VARCHAR),'-') AS [Quality],
                        ISNULL(CAST([Achievement] AS VARCHAR),'-') AS [Achievement],
                        CONVERT(VARCHAR,[Date],23) AS [Date]
                    FROM [dbo].[MachineEfficiency]
                    WHERE MONTH([Date])=@Month AND YEAR([Date])=@Year
                    ORDER BY [MachineName],[Date]";
                using var cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@Month", month);
                cmd.Parameters.AddWithValue("@Year", year);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Add(new
                    {
                        machineName = reader["MachineName"]?.ToString() ?? "-",
                        oEE = reader["OEE"]?.ToString() ?? "-",
                        operatingRatio = reader["OperatingRatio"]?.ToString() ?? "-",
                        ability = reader["Ability"]?.ToString() ?? "-",
                        quality = reader["Quality"]?.ToString() ?? "-",
                        achievement = reader["Achievement"]?.ToString() ?? "-",
                        date = reader["Date"]?.ToString() ?? "-"
                    });
                return new JsonResult(result);
            }
            catch (Exception ex) { return new JsonResult(new { error = ex.Message }); }
        }

        public class MachineEfficiencyInput
        {
            public string? MachineName { get; set; }
            public DateTime Date { get; set; }
            public string? Shift { get; set; }
            public double? PlanQty { get; set; }
            public double? DefectQty { get; set; }
            public double? GoodProductionQty { get; set; }
            public double? WorkingTime { get; set; }
            public double? QualityTrouble { get; set; }
            public double? ModelChangingLoss { get; set; }
            public double? MaterialShortageExternal { get; set; }
            public double? MachineToolsTrouble { get; set; }
            public double? ManPowerAdjustment { get; set; }
            public double? MaterialShortageInhouse { get; set; }
            public double? MaterialShortageInternal { get; set; }
            public double? SetRepairingLoss { get; set; }
            public double? GawseExternalBodies { get; set; }
            public double? Rework { get; set; }
            public double? MoldChangingLoss { get; set; }
            public double? BreakTime { get; set; }
            public double? CompanyActivity { get; set; }
            public double? MorningAssembly { get; set; }
            public double? Cleaning { get; set; }
            public double? StockOpname { get; set; }
            public double? GeneralAssembly { get; set; }
            public double? Maintenance { get; set; }
            public double? TrialRun { get; set; }
            public double? TrainingEducation { get; set; }
            public double? FreeTalkingQC { get; set; }
            public double? NoProductionDay { get; set; }
        }
    }
}