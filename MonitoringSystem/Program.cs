using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using MonitoringSystem.Hubs;
using MonitoringSystem.Data;
using MonitoringSystem.Filters;
// HAPUS: using static NuGet.Packaging.PackagingConstants; // tidak dipakai

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddDbContext<ApplicationDbContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));

builder.Services.AddDefaultIdentity<ApplicationUser>(
    options => options.SignIn.RequireConfirmedAccount = true)
    .AddEntityFrameworkStores<ApplicationDbContext>();

builder.Services.AddControllers();
builder.Services.AddRazorPages()
    .AddMvcOptions(options =>
    {
        options.Filters.Add<AuthorizeFilter>();
    });

builder.Services.AddHostedService<PlanUpdaterService>()
    .Configure<HostOptions>(options =>
    {
        options.BackgroundServiceExceptionBehavior = BackgroundServiceExceptionBehavior.Ignore;
    });

builder.Services.AddSignalR();

builder.Services.Configure<CookiePolicyOptions>(options =>
{
    options.CheckConsentNeeded = context => true;
    options.MinimumSameSitePolicy = SameSiteMode.None;
});

builder.Services.ConfigureApplicationCookie(options =>
{
    options.LoginPath = "/Identity/Account/Login";
    options.LogoutPath = "/Identity/Account/Logout";
    options.AccessDeniedPath = "/Identity/Account/AccessDenied";
    options.SlidingExpiration = true;
    options.ExpireTimeSpan = TimeSpan.FromHours(12);
});

// ✅ FIX 1: Compression HANYA di Production (dev konflik dengan BrowserLink → penyebab loading lama!)
if (!builder.Environment.IsDevelopment())
{
    builder.Services.AddResponseCompression(options =>
    {
        options.EnableForHttps = true;
    });
}

var app = builder.Build();

// ✅ FIX 2: Warm up database (tetap dipertahankan, bagus)
using (var scope = app.Services.CreateScope())
{
    var db = scope.ServiceProvider.GetRequiredService<ApplicationDbContext>();
    await db.Database.ExecuteSqlRawAsync("SELECT 1");
}

if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Error");
    app.UseHsts();
    app.UseResponseCompression(); // ✅ FIX 3: Pindah ke sini, hanya aktif di Production
}

app.UseHttpsRedirection();

// ✅ FIX 4: Cache static files tetap dipertahankan
app.UseStaticFiles(new StaticFileOptions
{
    OnPrepareResponse = ctx =>
    {
        ctx.Context.Response.Headers["Cache-Control"] = "public,max-age=604800";
    }
});

app.UseRouting();
app.UseCookiePolicy(); // ✅ FIX 5: TAMBAH INI — CookiePolicy harus di-use, bukan hanya di-configure!
app.UseAuthentication();
app.UseAuthorization();

app.MapHub<LossTimeHub>("/dataHub");
app.MapControllers();
app.MapRazorPages();

app.Run();