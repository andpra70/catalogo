const PROFILE_WIDGET_SCRIPT_ID = "ecosystem-profile-widget-script";
const PROFILE_WIDGET_SCRIPT_URL = "/auth/profile-widget.js";

function loadProfileWidget() {
  if (window.ProfileWidget) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const existing = document.getElementById(PROFILE_WIDGET_SCRIPT_ID);
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", reject, { once: true });
      return;
    }
    const script = document.createElement("script");
    script.id = PROFILE_WIDGET_SCRIPT_ID;
    script.src = PROFILE_WIDGET_SCRIPT_URL;
    script.async = true;
    script.addEventListener("load", resolve, { once: true });
    script.addEventListener("error", () => reject(new Error("Profile Widget non disponibile")), { once: true });
    document.head.appendChild(script);
  });
}

export async function openProfileWidget() {
  await loadProfileWidget();
  window.ProfileWidget.mount({ apiBase: "/auth", authWidgetUrl: "/auth/widget.js" });
  await new Promise((resolve) => requestAnimationFrame(resolve));
  await new Promise((resolve) => requestAnimationFrame(resolve));
  window.ProfileWidget.open();
}
