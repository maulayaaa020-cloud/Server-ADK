function toggleTheme() {
    var html = document.documentElement;
    var current = html.getAttribute('data-theme') || 'light';
    var next = current === 'dark' ? 'light' : 'dark';
    html.setAttribute('data-theme', next);
    localStorage.setItem('adkTheme', next);
    updateThemeIcons();
}

function updateThemeIcons() {
    var theme = document.documentElement.getAttribute('data-theme') || 'light';
    var icon  = theme === 'dark' ? '☀️' : '🌙';
    var title = theme === 'dark' ? 'Mode Siang' : 'Mode Malam';
    document.querySelectorAll('.theme-toggle').forEach(function(btn) {
        btn.textContent = icon;
        btn.title = title;
    });
}

document.addEventListener('DOMContentLoaded', updateThemeIcons);
