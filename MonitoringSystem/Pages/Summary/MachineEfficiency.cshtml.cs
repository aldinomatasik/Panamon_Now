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

        // Loss category definitions (single source of truth)
        private static readonly List<(string Col, string Group)> LossColumns = new()
        {
            ("QualityTrouble",           "WorkingLoss"),
            ("ModelChangingLoss",        "WorkingLoss"),
            ("MaterialShortageExternal", "WorkingLoss"),
            ("MachineToolsTrouble",      "WorkingLoss"),
            ("ManPowerAdjustment",       "WorkingLoss"),
            ("MaterialShortageInhouse",  "WorkingLoss"),
            ("MaterialShortageInternal", "WorkingLoss"),
            ("SetRepairingLoss",         "WorkingLoss"),
            ("GawseExternalBodies",      "WorkingLoss"),
            ("Rework",                   "WorkingLoss"),
            ("MoldChangingLoss",         "WorkingLoss"),
            ("BreakTime",                "FixedLoss"),
            ("CompanyActivity",          "FixedLoss"),
            ("MorningAssembly",          "FixedLoss"),
            ("Cleaning",                 "FixedLoss"),
            ("StockOpname",              "FixedLoss"),
            ("GeneralAssembly",          "FixedLoss"),
            ("Maintenance",              "FixedLoss"),
            ("TrialRun",                 "FixedLoss"),
            ("TrainingEducation",        "FixedLoss"),
            ("FreeTalkingQC",            "FixedLoss"),
            ("NoProductionDay",          "FixedLoss"),
        };

        public MachineEfficiencyModel(IConfiguration config, IWebHostEnvironment env)
        {
            _config = config;
            _env = env;
        }

        public void OnGet() { }

        // =====================================================
        // HELPER: Normalisasi shift format
        // =====================================================
        private static string NormalizeShift(string? shift) => (shift ?? "") switch
        {
            "1" or "Shift 1" => "Shift 1",
            "2" or "Shift 2" => "Shift 2",
            "3" or "Shift 3" => "Shift 3",
            "NS" or "Non Shift" => "Non Shift",
            var s => s
        };

        // =====================================================
        // HELPER: Hitung semua metric dari input
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
                double totalLoss = input.LossItems?.Sum(x => x.LossMinutes ?? 0) ?? 0;
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
        // API: ?handler=Machines
        // =====================================================
        public JsonResult OnGetMachines()
        {
            var machines = new List<object>();
            try
            {
                using var conn = OpenConn();
                using var cmd = new SqlCommand("SELECT DISTINCT MachineName FROM [dbo].[MachineEfficiency] WHERE MachineName IS NOT NULL ORDER BY MachineName", conn);
                using var reader = cmd.ExecuteReader();
                // BENAR - value pakai MachineName karena tidak ada kolom ID
                while (reader.Read())
                    machines.Add(new { value = reader["MachineName"].ToString(), label = reader["MachineName"].ToString() });
            }
            catch (Exception ex) { return new JsonResult(new { error = ex.Message }); }
            return new JsonResult(machines);
        }

        // =====================================================
        // API: ?handler=LossCategories
        // =====================================================

        public JsonResult OnGetLossCategories()
        {
            static string ToLabel(string col) =>
                System.Text.RegularExpressions.Regex.Replace(col, @"(?<=[a-z])(?=[A-Z])|(?<=[A-Z])(?=[A-Z][a-z])", " ");

            var workingLoss = LossColumns
                .Where(x => x.Group == "WorkingLoss")
                .Select(x => new { value = x.Col, label = ToLabel(x.Col) });

            var fixedLoss = LossColumns
                .Where(x => x.Group == "FixedLoss")
                .Select(x => new { value = x.Col, label = ToLabel(x.Col) });

            return new JsonResult(new { workingLoss, fixedLoss });
        }

        // =====================================================
        // API: ?handler=WorkingTime&machineName=X&date=Y&shift=Z
        // =====================================================
        public JsonResult OnGetWorkingTime(string machineName, DateTime date, string shift)
        {
            try
            {
                string shiftNorm = NormalizeShift(shift);
                using var conn = OpenConn();
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
                cmd.Parameters.AddWithValue("@Shift", shiftNorm);
                var result = cmd.ExecuteScalar();
                if (result != null && result != DBNull.Value)
                    return new JsonResult(new { found = true, workingTime = Convert.ToDouble(result) });
                return new JsonResult(new { found = false });
            }
            catch (Exception ex) { return new JsonResult(new { found = false, error = ex.Message }); }
        }

        // =====================================================
        // API: ?handler=LoadExisting&machineName=X&date=Y&shift=Z
        // =====================================================
        public JsonResult OnGetLoadExisting(string machineName, DateTime date, string shift)
        {
            try
            {
                string shiftNorm = NormalizeShift(shift);
                using var conn = OpenConn();

                // Ambil header
                string sqlHeader = @"
                    SELECT TOP 1 ID, PlanQty, DefectQty, GoodProductionQty, WorkingTime
                    FROM [dbo].[MachineEfficiency]
                    WHERE MachineName = @MachineName
                      AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                      AND Shift = @Shift
                    ORDER BY ID DESC";
                using var cmd = new SqlCommand(sqlHeader, conn);
                cmd.Parameters.AddWithValue("@MachineName", machineName ?? "");
                cmd.Parameters.AddWithValue("@Date", date);
                cmd.Parameters.AddWithValue("@Shift", shiftNorm);
                using var reader = cmd.ExecuteReader();

                if (!reader.Read()) return new JsonResult(new { found = false });

                int efficiencyId = Convert.ToInt32(reader["ID"]);
                double? G(string col) => reader[col] == DBNull.Value ? null : Convert.ToDouble(reader[col]);
                double? planQty = G("PlanQty");
                double? defectQty = G("DefectQty");
                double? goodProductionQty = G("GoodProductionQty");
                double? workingTime = G("WorkingTime");
                reader.Close();

                // Ambil loss items
                string sqlLoss = @"
                    SELECT LossCategory, LossGroup, LossMinutes
                    FROM [dbo].[MachineEfficiencyLoss]
                    WHERE EfficiencyID = @EfficiencyID";
                using var cmdLoss = new SqlCommand(sqlLoss, conn);
                cmdLoss.Parameters.AddWithValue("@EfficiencyID", efficiencyId);
                using var readerLoss = cmdLoss.ExecuteReader();

                var lossDict = new Dictionary<string, double?>();
                while (readerLoss.Read())
                {
                    string cat = readerLoss["LossCategory"]?.ToString() ?? "";
                    double? min = readerLoss["LossMinutes"] == DBNull.Value ? null : Convert.ToDouble(readerLoss["LossMinutes"]);
                    lossDict[cat] = min;
                }

                // Build response compatible dengan frontend lama
                return new JsonResult(new
                {
                    found = true,
                    planQty,
                    defectQty,
                    goodProductionQty,
                    workingTime,
                    qualityTrouble = lossDict.GetValueOrDefault("QualityTrouble"),
                    modelChangingLoss = lossDict.GetValueOrDefault("ModelChangingLoss"),
                    materialShortageExternal = lossDict.GetValueOrDefault("MaterialShortageExternal"),
                    machineToolsTrouble = lossDict.GetValueOrDefault("MachineToolsTrouble"),
                    manPowerAdjustment = lossDict.GetValueOrDefault("ManPowerAdjustment"),
                    materialShortageInhouse = lossDict.GetValueOrDefault("MaterialShortageInhouse"),
                    materialShortageInternal = lossDict.GetValueOrDefault("MaterialShortageInternal"),
                    setRepairingLoss = lossDict.GetValueOrDefault("SetRepairingLoss"),
                    gawseExternalBodies = lossDict.GetValueOrDefault("GawseExternalBodies"),
                    rework = lossDict.GetValueOrDefault("Rework"),
                    moldChangingLoss = lossDict.GetValueOrDefault("MoldChangingLoss"),
                    breakTime = lossDict.GetValueOrDefault("BreakTime"),
                    companyActivity = lossDict.GetValueOrDefault("CompanyActivity"),
                    morningAssembly = lossDict.GetValueOrDefault("MorningAssembly"),
                    cleaning = lossDict.GetValueOrDefault("Cleaning"),
                    stockOpname = lossDict.GetValueOrDefault("StockOpname"),
                    generalAssembly = lossDict.GetValueOrDefault("GeneralAssembly"),
                    maintenance = lossDict.GetValueOrDefault("Maintenance"),
                    trialRun = lossDict.GetValueOrDefault("TrialRun"),
                    trainingEducation = lossDict.GetValueOrDefault("TrainingEducation"),
                    freeTalkingQC = lossDict.GetValueOrDefault("FreeTalkingQC"),
                    noProductionDay = lossDict.GetValueOrDefault("NoProductionDay"),
                });
            }
            catch (Exception ex) { return new JsonResult(new { found = false, error = ex.Message }); }
        }

        // =====================================================
        // API: ?handler=Submit (UPSERT)
        // =====================================================
        public JsonResult OnPostSubmit([FromBody] MachineEfficiencyInput input)
        {
            try
            {
                input.Shift = NormalizeShift(input.Shift);
                var (quality, operatingRatio, ability, oee, achievement, category) = CalcMetrics(input);

                using var conn = OpenConn();
                using var tx = conn.BeginTransaction();

                // Cek existing
                int? existingId = null;
                using (var checkCmd = new SqlCommand(@"
                    SELECT TOP 1 ID FROM [dbo].[MachineEfficiency]
                    WHERE MachineName = @MachineName
                      AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
                      AND Shift = @Shift", conn, tx))
                {
                    checkCmd.Parameters.AddWithValue("@MachineName", input.MachineName ?? "");
                    checkCmd.Parameters.AddWithValue("@Date", input.Date);
                    checkCmd.Parameters.AddWithValue("@Shift", input.Shift ?? "");
                    var result = checkCmd.ExecuteScalar();
                    if (result != null && result != DBNull.Value)
                        existingId = Convert.ToInt32(result);
                }

                int efficiencyId;
                if (existingId.HasValue)
                {
                    // UPDATE header
                    using var updateCmd = new SqlCommand(@"
                        UPDATE [dbo].[MachineEfficiency] SET
                            WorkingTime = @WorkingTime, PlanQty = @PlanQty,
                            GoodProductionQty = @GoodProductionQty, DefectQty = @DefectQty,
                            OEE = @OEE, OperatingRatio = @OperatingRatio,
                            Ability = @Ability, Quality = @Quality,
                            Achievement = @Achievement, Category = @Category
                        WHERE ID = @ID", conn, tx);
                    AddHeaderParams(updateCmd, input, quality, operatingRatio, ability, oee, achievement, category);
                    updateCmd.Parameters.AddWithValue("@ID", existingId.Value);
                    updateCmd.ExecuteNonQuery();

                    // Hapus loss lama
                    using var delLoss = new SqlCommand("DELETE FROM [dbo].[MachineEfficiencyLoss] WHERE EfficiencyID = @ID", conn, tx);
                    delLoss.Parameters.AddWithValue("@ID", existingId.Value);
                    delLoss.ExecuteNonQuery();

                    efficiencyId = existingId.Value;
                }
                else
                {
                    // INSERT header
                    using var insertCmd = new SqlCommand(@"
                        INSERT INTO [dbo].[MachineEfficiency]
                            (MachineName, Date, Shift, WorkingTime, PlanQty, GoodProductionQty, DefectQty,
                             OEE, OperatingRatio, Ability, Quality, Achievement, Category)
                        VALUES
                            (@MachineName, @Date, @Shift, @WorkingTime, @PlanQty, @GoodProductionQty, @DefectQty,
                             @OEE, @OperatingRatio, @Ability, @Quality, @Achievement, @Category);
                        SELECT SCOPE_IDENTITY();", conn, tx);
                    AddHeaderParams(insertCmd, input, quality, operatingRatio, ability, oee, achievement, category);
                    efficiencyId = Convert.ToInt32(insertCmd.ExecuteScalar());
                }

                // INSERT loss items
                InsertLossItems(conn, tx, efficiencyId, input.LossItems);

                tx.Commit();
                return new JsonResult(new { success = true, action = existingId.HasValue ? "updated" : "inserted" });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, error = ex.Message });
            }
        }

        // =====================================================
        // API: ?handler=ImportExcel
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
                        double? D(int row) => TryParseDouble(sheet.Cells[row, col].Value);

                        var lossItems = new List<LossItem>();
                        void AddLoss(string cat, string grp, double? min) { if (min.HasValue) lossItems.Add(new LossItem { LossCategory = cat, LossGroup = grp, LossMinutes = min }); }

                        AddLoss("QualityTrouble", "WorkingLoss", D(ROW_QUALITY_TROUBLE));
                        AddLoss("ModelChangingLoss", "WorkingLoss", D(ROW_MODEL_CHANGING_LOSS));
                        AddLoss("MaterialShortageExternal", "WorkingLoss", D(ROW_MATERIAL_SHORTAGE_EXT));
                        AddLoss("MachineToolsTrouble", "WorkingLoss", D(ROW_MACHINE_TOOLS_TROUBLE));
                        AddLoss("ManPowerAdjustment", "WorkingLoss", D(ROW_MAN_POWER_ADJUSTMENT));
                        AddLoss("MaterialShortageInhouse", "WorkingLoss", D(ROW_MATERIAL_SHORTAGE_INHOUSE));
                        AddLoss("MaterialShortageInternal", "WorkingLoss", D(ROW_MATERIAL_SHORTAGE_INTERNAL));
                        AddLoss("SetRepairingLoss", "WorkingLoss", D(ROW_SET_REPAIRING_LOSS));
                        AddLoss("GawseExternalBodies", "WorkingLoss", D(ROW_GAWSE_EXTERNAL_BODIES));
                        AddLoss("Rework", "WorkingLoss", D(ROW_REWORK));
                        AddLoss("MoldChangingLoss", "WorkingLoss", D(ROW_MOLD_CHANGING_LOSS));
                        AddLoss("BreakTime", "FixedLoss", D(ROW_BREAK_TIME));
                        AddLoss("CompanyActivity", "FixedLoss", D(ROW_COMPANY_ACTIVITY));
                        AddLoss("MorningAssembly", "FixedLoss", D(ROW_MORNING_ASSEMBLY));
                        AddLoss("Cleaning", "FixedLoss", D(ROW_CLEANING));
                        AddLoss("StockOpname", "FixedLoss", D(ROW_STOCK_OPNAME));
                        AddLoss("GeneralAssembly", "FixedLoss", D(ROW_GENERAL_ASSEMBLY));
                        AddLoss("Maintenance", "FixedLoss", D(ROW_MAINTENANCE));
                        AddLoss("TrialRun", "FixedLoss", D(ROW_TRIAL_RUN));
                        AddLoss("TrainingEducation", "FixedLoss", D(ROW_TRAINING_EDUCATION));
                        AddLoss("FreeTalkingQC", "FixedLoss", D(ROW_FREE_TALKING_QC));
                        AddLoss("NoProductionDay", "FixedLoss", D(ROW_NO_PRODUCTION_DAY));

                        double? planQty = D(ROW_PLAN);
                        double? goodQty = D(ROW_GOOD_PRODUCTION_QTY);
                        double? defectQty = D(ROW_DEFECT_QTY);
                        double? workingTime = D(ROW_WORKING_TIME);

                        bool hasAnyData = planQty.HasValue || goodQty.HasValue || defectQty.HasValue
                            || workingTime.HasValue || lossItems.Count > 0;
                        if (!hasAnyData) continue;

                        newData.Add(new MachineEfficiencyInput
                        {
                            MachineName = machineName,
                            Date = new DateTime(year, month, day),
                            Shift = shifts[shiftIndex],
                            PlanQty = planQty,
                            GoodProductionQty = goodQty,
                            DefectQty = defectQty,
                            WorkingTime = workingTime,
                            LossItems = lossItems
                        });
                    }
                }

                // ✅ BARU — hapus hanya Date+Shift yang ada di Excel
                using var conn = OpenConn();
                using var tx = conn.BeginTransaction();

                foreach (var item in newData)
                {
                    // Cari existing ID untuk Date + Shift ini
                    int? existingId = null;
                    using (var checkCmd = new SqlCommand(@"
        SELECT TOP 1 ID FROM [dbo].[MachineEfficiency]
        WHERE MachineName = @MachineName
          AND CAST([Date] AS DATE) = CAST(@Date AS DATE)
          AND Shift = @Shift", conn, tx))
                    {
                        checkCmd.Parameters.AddWithValue("@MachineName", machineName);
                        checkCmd.Parameters.AddWithValue("@Date", item.Date);
                        checkCmd.Parameters.AddWithValue("@Shift", item.Shift ?? "");
                        var result = checkCmd.ExecuteScalar();
                        if (result != null && result != DBNull.Value)
                            existingId = Convert.ToInt32(result);
                    }

                    if (existingId.HasValue)
                    {
                        // Hapus loss lama untuk hari+shift ini saja
                        using var delLoss = new SqlCommand(@"
            DELETE FROM [dbo].[MachineEfficiencyLoss]
            WHERE EfficiencyID = @ID", conn, tx);
                        delLoss.Parameters.AddWithValue("@ID", existingId.Value);
                        delLoss.ExecuteNonQuery();

                        // Hapus header lama untuk hari+shift ini saja
                        using var delHeader = new SqlCommand(@"
            DELETE FROM [dbo].[MachineEfficiency]
            WHERE ID = @ID", conn, tx);
                        delHeader.Parameters.AddWithValue("@ID", existingId.Value);
                        delHeader.ExecuteNonQuery();

                        Console.WriteLine($"[ImportExcel] Replaced: {machineName} {item.Date:yyyy-MM-dd} {item.Shift}");
                    }
                }

                int insertedCount = 0;
                foreach (var item in newData)
                {
                    var (quality, operatingRatio, ability, oee, achievement, category) = CalcMetrics(item);

                    using var insertCmd = new SqlCommand(@"
                        INSERT INTO [dbo].[MachineEfficiency]
                            (MachineName, Date, Shift, WorkingTime, PlanQty, GoodProductionQty, DefectQty,
                             OEE, OperatingRatio, Ability, Quality, Achievement, Category)
                        VALUES
                            (@MachineName, @Date, @Shift, @WorkingTime, @PlanQty, @GoodProductionQty, @DefectQty,
                             @OEE, @OperatingRatio, @Ability, @Quality, @Achievement, @Category);
                        SELECT SCOPE_IDENTITY();", conn, tx);
                    AddHeaderParams(insertCmd, item, quality, operatingRatio, ability, oee, achievement, category);
                    int efficiencyId = Convert.ToInt32(insertCmd.ExecuteScalar());

                    InsertLossItems(conn, tx, efficiencyId, item.LossItems);
                    insertedCount++;
                }

                tx.Commit();
                return new JsonResult(new { success = true, message = $"Berhasil import {insertedCount} data untuk {machineName} (Bulan {month}/{year})." });
            }
            catch (Exception ex)
            {
                return new JsonResult(new { success = false, error = "Gagal Import: " + ex.Message });
            }
        }

        // =====================================================
        // API: ?handler=MachineList
        // =====================================================
        public JsonResult OnGetMachineList()
        {
            var result = new List<object>();
            try
            {
                using var conn = OpenConn();
                using var cmd = new SqlCommand(@"
                    SELECT DISTINCT MachineName FROM [dbo].[MachineEfficiency]
                    WHERE MachineName IS NOT NULL AND MachineName != ''
                    ORDER BY MachineName", conn);
                using var reader = cmd.ExecuteReader();
                while (reader.Read())
                    result.Add(new { machineName = reader["MachineName"]?.ToString() ?? "-" });
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
                using var conn = OpenConn();
                string sql = @"
                    SELECT MachineName,
                        ISNULL(CAST(OEE AS VARCHAR),'-') AS OEE,
                        ISNULL(CAST(OperatingRatio AS VARCHAR),'-') AS OperatingRatio,
                        ISNULL(CAST(Ability AS VARCHAR),'-') AS Ability,
                        ISNULL(CAST(Quality AS VARCHAR),'-') AS Quality,
                        ISNULL(CAST(Achievement AS VARCHAR),'-') AS Achievement,
                        CONVERT(VARCHAR, Date, 23) AS Date
                    FROM [dbo].[MachineEfficiency]
                    WHERE MONTH(Date) = @Month AND YEAR(Date) = @Year
                    ORDER BY MachineName, Date";
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

        // =====================================================
        // HELPER: Insert loss items ke MachineEfficiencyLoss
        // =====================================================
        private static void InsertLossItems(SqlConnection conn, SqlTransaction tx, int efficiencyId, List<LossItem>? lossItems)
        {
            if (lossItems == null || lossItems.Count == 0) return;
            foreach (var loss in lossItems.Where(l => l.LossMinutes.HasValue))
            {
                using var cmd = new SqlCommand(@"
                    INSERT INTO [dbo].[MachineEfficiencyLoss] (EfficiencyID, LossCategory, LossGroup, LossMinutes)
                    VALUES (@EfficiencyID, @LossCategory, @LossGroup, @LossMinutes)", conn, tx);
                cmd.Parameters.AddWithValue("@EfficiencyID", efficiencyId);
                cmd.Parameters.AddWithValue("@LossCategory", loss.LossCategory ?? "");
                cmd.Parameters.AddWithValue("@LossGroup", loss.LossGroup ?? "");
                cmd.Parameters.AddWithValue("@LossMinutes", (object?)loss.LossMinutes ?? DBNull.Value);
                cmd.ExecuteNonQuery();
            }
        }

        // =====================================================
        // HELPER: Tambah header params ke SqlCommand
        // =====================================================
        private static void AddHeaderParams(SqlCommand cmd, MachineEfficiencyInput input,
            double? quality, double? operatingRatio, double? ability, double? oee,
            double? achievement, string? category)
        {
            cmd.Parameters.AddWithValue("@MachineName", input.MachineName ?? "");
            cmd.Parameters.AddWithValue("@Date", input.Date);
            cmd.Parameters.AddWithValue("@Shift", input.Shift ?? "");
            cmd.Parameters.AddWithValue("@WorkingTime", (object?)input.WorkingTime ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@PlanQty", (object?)input.PlanQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@GoodProductionQty", (object?)input.GoodProductionQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@DefectQty", (object?)input.DefectQty ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OEE", (object?)oee ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@OperatingRatio", (object?)operatingRatio ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Ability", (object?)ability ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Quality", (object?)quality ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Achievement", (object?)achievement ?? DBNull.Value);
            cmd.Parameters.AddWithValue("@Category", (object?)category ?? DBNull.Value);
        }

        // =====================================================
        // HELPER: Open DB connection
        // =====================================================
        private SqlConnection OpenConn()
        {
            var conn = new SqlConnection(_config.GetConnectionString("MachineConnection")!);
            conn.Open();
            return conn;
        }

        // =====================================================
        // HELPER: Parse double dari Excel cell
        // =====================================================
        private static double? TryParseDouble(object? val)
        {
            if (val == null) return null;
            return double.TryParse(val.ToString(), out double result) ? result : null;
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
        // Models
        // =====================================================
        public class LossItem
        {
            public string? LossCategory { get; set; }
            public string? LossGroup { get; set; }
            public double? LossMinutes { get; set; }
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
            public List<LossItem>? LossItems { get; set; }
        }
    }
}