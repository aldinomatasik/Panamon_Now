/**
 * Layout.js - MonitoringSystem
 * Menggabungkan semua functionality dari _Layout.cshtml
 */

(function ($) {
    'use strict';

    /**
     * 1. ACTIVE NAVIGATION HIGHLIGHT
     * Menandai menu yang aktif berdasarkan URL saat ini
     */
    function initActiveNavigation() {
        var currentPath = window.location.pathname.toLowerCase();

        if (currentPath.endsWith('/index')) {
            currentPath = currentPath.substring(0, currentPath.length - 6);
        }
        if (currentPath === '') {
            currentPath = '/';
        }

        $('#bdSidebar .nav-item a').removeClass('active');

        $('#bdSidebar .nav-item a').each(function () {
            var $this = $(this);
            var href = $this.attr('href');

            if (!href || href === '#') return;

            var linkPath = href.replace('~', '').toLowerCase();

            if (linkPath.endsWith('/index')) {
                linkPath = linkPath.substring(0, linkPath.length - 6);
            }
            if (linkPath === '') {
                linkPath = '/';
            }

            if (currentPath === linkPath) {
                $this.addClass('active');
            } else if (currentPath.startsWith(linkPath + '/')) {
                $this.addClass('active');
            }
        });
    }

    /**
     * 2. DATE & TIME UPDATE
     * Update tanggal dan waktu real-time
     */
    function updateDateTime() {
        const now = new Date();
        const dateStr = now.toLocaleDateString('id-ID', {
            day: '2-digit',
            month: '2-digit',
            year: 'numeric'
        }).replace(/\//g, '.');

        const timeStr = now.toLocaleTimeString('id-ID', {
            hour: '2-digit',
            minute: '2-digit',
            second: '2-digit',
            hour12: false
        });

        document.getElementById('current-date').textContent = dateStr;
        document.getElementById('current-time').textContent = timeStr;
    }

    function initDateTime() {
        updateDateTime();
        setInterval(updateDateTime, 1000);
    }

    /**
     * 3. MORE MENU TOGGLE
     */
    function initMoreMenu() {
        const moreBtn = document.getElementById('moreMenuBtn');
        const morePopup = document.getElementById('moreMenuPopup');

        if (moreBtn && morePopup) {
            moreBtn.addEventListener('click', function (e) {
                e.preventDefault();
                e.stopPropagation();

                if (morePopup.style.display === 'none' || morePopup.style.display === '') {
                    morePopup.style.display = 'block';
                } else {
                    morePopup.style.display = 'none';
                }
            });

            document.addEventListener('click', function (e) {
                if (!moreBtn.contains(e.target) && !morePopup.contains(e.target)) {
                    morePopup.style.display = 'none';
                }
            });
        }
    }

    /**
     * 4. THEME TOGGLE (DARK / LIGHT MODE)
     */
    function getCookie(name) {
        return document.cookie.split('; ')
            .find(r => r.startsWith(name + '='))?.split('=')[1];
    }

    /**
     * 5. UPDATE LOGO & JAM SESUAI LIGHT MODE
     */
    function updateLogoAndJam() {
        const isLightMode = document.body.classList.contains('light-mode');
        const logoImg = document.querySelector('.header-left img');
        const currentTime = document.getElementById('current-time');
        const currentDate = document.getElementById('current-date');
        const headerCenter = document.querySelector('.header-center');

        if (isLightMode) {
            // 🌅 LIGHT MODE
            if (logoImg) logoImg.src = '/assets/PanasonicIconBlack.png';
            if (currentTime) currentTime.style.color = '#1a1a2e';
            if (currentDate) currentDate.style.color = '#1a1a2e';
            if (headerCenter) headerCenter.style.color = '#1a1a2e';
        } else {
            // 🌙 DARK MODE
            if (logoImg) logoImg.src = '/assets/PanasonicIcon.png';
            if (currentTime) currentTime.style.color = '#ffffff';
            if (currentDate) currentDate.style.color = '#ffffff';
            if (headerCenter) headerCenter.style.color = '#ffffff';
        }
    }

    function applyTheme(theme) {
        const themeIcon = document.getElementById('themeIcon');

        if (theme === 'light') {
            document.body.classList.add('light-mode');
            document.documentElement.classList.add('light-mode');
            if (themeIcon) themeIcon.className = 'fa-solid fa-moon';
        } else {
            document.body.classList.remove('light-mode');
            document.documentElement.classList.remove('light-mode');
            if (themeIcon) themeIcon.className = 'fa-solid fa-sun';
        }

        // Update logo & jam
        updateLogoAndJam();

        // ✅ Update semua chart langsung saat theme di-apply
        setTimeout(updateAllChartsTheme, 80);
    }

    function initThemeToggle() {
        const themeToggle = document.getElementById('themeToggle');

        // Load saved theme dari cookie, default dark
        const savedTheme = getCookie('themeMode') || 'dark';
        applyTheme(savedTheme);

        if (themeToggle) {
            themeToggle.addEventListener('click', function () {
                const isLight = document.body.classList.contains('light-mode');
                const newTheme = isLight ? 'dark' : 'light';

                // Simpan ke cookie selama 1 tahun
                document.cookie = `themeMode=${newTheme}; path=/; max-age=31536000`;
                applyTheme(newTheme);
            });
        }
    }

    /**
     * INITIALIZATION
     */
    $(document).ready(function () {
        initActiveNavigation();
        initDateTime();
    });

    document.addEventListener('DOMContentLoaded', function () {
        initMoreMenu();
        initThemeToggle();
        updateLogoAndJam();

        // ✅ Jalankan sekali setelah semua chart selesai dibuat
        setTimeout(updateAllChartsTheme, 300);
    });

})(jQuery);


// ============================================================
// 6. UNIVERSAL THEME-AWARE CHART UPDATER
// Otomatis berlaku ke SEMUA page yang punya Chart.js
// Tidak perlu tambah kode apapun di masing-masing page
// ============================================================

(function () {

    // ---- Helper: Deteksi mode saat ini ----
    const isLightModeActive = () =>
        document.documentElement.classList.contains('light-mode') ||
        document.body.classList.contains('light-mode');

    // ---- Helper: Ambil warna sesuai mode ----
    const getThemeColors = () => {
        const isLight = isLightModeActive();
        return {
            textColor: isLight ? '#1a1a2e' : '#ffffff',
            gridColor: isLight ? 'rgba(26, 26, 46, 0.12)' : 'rgba(255, 255, 255, 0.15)',
            tooltipBg: isLight ? 'rgba(240, 245, 255, 0.97)' : 'rgba(0, 0, 0, 0.82)',
            tooltipBorder: isLight ? 'rgba(26, 26, 46, 0.15)' : 'rgba(255, 255, 255, 0.15)'
        };
    };

    // ---- Fungsi utama: update SEMUA chart yang ada di halaman ----
    // Dibuat global (window) supaya bisa dipanggil dari applyTheme() di atas
    window.updateAllChartsTheme = function () {
        if (typeof Chart === 'undefined') return;

        const c = getThemeColors();

        // Chart.js v3+ simpan semua instance di Chart.instances
        const allCharts = Object.values(Chart.instances || {});

        allCharts.forEach(chart => {
            if (!chart || !chart.options) return;

            // ---- Legend ----
            if (chart.options.plugins?.legend?.labels) {
                chart.options.plugins.legend.labels.color = c.textColor;
            }

            // ---- Tooltip ----
            if (chart.options.plugins?.tooltip) {
                chart.options.plugins.tooltip.titleColor = c.textColor;
                chart.options.plugins.tooltip.bodyColor = c.textColor;
                chart.options.plugins.tooltip.backgroundColor = c.tooltipBg;
                chart.options.plugins.tooltip.borderColor = c.tooltipBorder;
            }

            // ---- Semua Scales (x, y, y1, y_ratio, dll) ----
            Object.values(chart.options.scales || {}).forEach(scale => {
                if (!scale) return;
                if (scale.ticks) scale.ticks.color = c.textColor;
                if (scale.title) scale.title.color = c.textColor;
                if (scale.grid) scale.grid.color = c.gridColor;
            });

            // Apply tanpa animasi biar instant
            chart.update('none');
        });
    };

    // ---- MutationObserver: watch class change di <html> DAN <body> ----
    // Fallback kalau theme di-toggle dari cara apapun
    const observer = new MutationObserver(function (mutations) {
        mutations.forEach(function (mutation) {
            if (mutation.attributeName === 'class') {
                window.updateAllChartsTheme();
            }
        });
    });

    observer.observe(document.documentElement, { attributes: true });
    observer.observe(document.body, { attributes: true });

})();