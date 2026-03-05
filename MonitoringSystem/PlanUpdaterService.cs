using Microsoft.Extensions.Hosting;
using Microsoft.EntityFrameworkCore;
using System;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using MonitoringSystem.Data;
using MonitoringSystem.Models;

public class PlanUpdaterService : BackgroundService
{
    private readonly IServiceProvider _serviceProvider;

    public PlanUpdaterService(IServiceProvider serviceProvider)
    {
        _serviceProvider = serviceProvider;
    }

    protected override async Task ExecuteAsync(CancellationToken stoppingToken)
    {
        while (!stoppingToken.IsCancellationRequested)
        {
            try
            {
                await Task.Delay(7000, stoppingToken);  // Delay 7 detik

                // Panggil metode untuk mengganti data setiap interval waktu
                UpdatePlanData();
            }
            catch (Exception ex)
            {
                // Tangani exception di sini dan log untuk troubleshooting
                System.Diagnostics.Debug.WriteLine($"Error occurred: {ex.Message}");
                // Optionally log to a logging service if needed
            }
        }
    }


    private void UpdatePlanData()
    {
        var currentTime = DateTime.Now;

        // Tentukan rentang waktu untuk penggantian data
        if (currentTime.Hour >= 7 && currentTime.Hour < 16)
        {
            // Data untuk 7 AM - 4 PM
            InsertOrReplacePlan("07:00", "16:00", "MCH1-01");  // CU
            InsertOrReplacePlan("07:00", "16:00", "MCH1-02");  // CS
        }
        else if (currentTime.Hour >= 16 && currentTime.Hour < 23)
        {
            // Data untuk 4 PM - 11:45 PM
            InsertOrReplacePlan("16:00", "23:45", "MCH1-01");  // CU
            InsertOrReplacePlan("16:00", "23:45", "MCH1-02");  // CS
        }
        else
        {
            // Data untuk 11:45 PM - 7 AM
            InsertOrReplacePlan("23:45", "07:00", "MCH1-01");  // CU
            InsertOrReplacePlan("23:45", "07:00", "MCH1-02");  // CS
        }
    }

    private void InsertOrReplacePlan(string startTime, string endTime, string machineCode)
    {
        // Implementasi logika untuk mengganti data berdasarkan waktu dan MachineCode
        using (var scope = _serviceProvider.CreateScope())
        {
            var context = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();

            //Query untuk replace atau insert data baru berdasarkan MachineCode
           var existingRecord = context.HourlyPlanData
               .FirstOrDefault(x => x.SelectedDate == DateTime.Today && x.MachineCode == machineCode);

            if (existingRecord != null)
            {
                // Replace data
                existingRecord.TotalPlan = CalculatePlan();
                existingRecord.UpdatedAt = DateTime.Now;
                context.Update(existingRecord);
            }
            else
            {
                // Insert data baru
                context.HourlyPlanData.Add(new HourlyPlanData
                {
                    MachineCode = machineCode, 
                    SelectedDate = DateTime.Today,
                    TotalPlan = CalculatePlan(),
                    CreatedAt = DateTime.Now,
                    UpdatedAt = DateTime.Now
                });
            }

            context.SaveChanges();
        }
    }

    private int CalculatePlan()
    {
        // Implementasi logika untuk menghitung jumlah plan
        return 10; // Ganti dengan logika yang sesuai untuk menghitung nilai plan
    }
}
