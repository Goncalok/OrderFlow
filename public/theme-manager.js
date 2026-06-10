const THEME_STORAGE_KEY = "orderflow-theme";

class ThemeManager {
  constructor() {
    this.currentTheme = "orderflow";
    this.init();
  }

  init() {
    this.applyTheme("orderflow");
    this.updateThemeSwitcherUI();
  }

  getStoredTheme() {
    return "orderflow";
  }

  saveTheme() {
    localStorage.setItem(THEME_STORAGE_KEY, "orderflow");
  }

  applyTheme() {
    document.documentElement.setAttribute("data-theme", "orderflow");
    document.body.setAttribute("data-theme", "orderflow");
    this.currentTheme = "orderflow";
  }

  ensureThemeCSSLoaded() {
    return;
  }

  switchTheme() {
    this.saveTheme();
    this.applyTheme();
    this.updateThemeSwitcherUI();
    return "orderflow";
  }

  getCurrentTheme() {
    return "orderflow";
  }

  isMinimalistModern() {
    return true;
  }

  observeThemeChanges() {
    return;
  }

  updateThemeSwitcherUI() {
    const themeToggle = document.getElementById("themeToggleButton");
    if (!themeToggle) return;
    themeToggle.classList.add("theme-active");
    themeToggle.setAttribute("data-theme-status", "orderflow");
  }
}

let themeManager;
if (document.readyState === "loading") {
  document.addEventListener("DOMContentLoaded", () => {
    themeManager = new ThemeManager();
  });
} else {
  themeManager = new ThemeManager();
}
