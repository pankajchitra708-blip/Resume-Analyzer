(function () {
    const storageKey = "theme";

    function preferredTheme() {
        const saved = localStorage.getItem(storageKey);
        if (saved === "light" || saved === "dark") return saved;
        return window.matchMedia && window.matchMedia("(prefers-color-scheme: light)").matches ? "light" : "dark";
    }

    function labelFor(theme) {
        return theme === "light" ? "Dark Mode" : "Light Mode";
    }

    function applyTheme(theme) {
        document.documentElement.setAttribute("data-theme", theme);
        localStorage.setItem(storageKey, theme);
        document.querySelectorAll(".theme-toggle").forEach((button) => {
            const label = labelFor(theme);
            button.innerHTML = '<span class="toggle-orbit" aria-hidden="true"><span class="toggle-icon"></span></span><span class="toggle-label">' + label + "</span>";
            button.setAttribute("aria-label", "Switch to " + label);
            button.setAttribute("aria-pressed", theme === "dark" ? "true" : "false");
            button.dataset.theme = theme;
        });
    }

    window.toggleTheme = function () {
        const current = document.documentElement.getAttribute("data-theme") || preferredTheme();
        applyTheme(current === "light" ? "dark" : "light");
    };

    function ensureToggle() {
        let button = document.getElementById("themeToggle");
        if (!button) {
            button = document.createElement("button");
            button.id = "themeToggle";
            button.type = "button";
            button.className = "theme-toggle";
            document.body.prepend(button);
        }
        button.classList.add("theme-toggle");
        button.type = "button";
        button.onclick = window.toggleTheme;
    }

    document.addEventListener("DOMContentLoaded", function () {
        ensureToggle();
        applyTheme(preferredTheme());
    });

    window.addEventListener("load", function () {
        ensureToggle();
        applyTheme(document.documentElement.getAttribute("data-theme") || preferredTheme());
    });
})();
