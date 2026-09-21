import { useEffect, useState } from "react";

const GDPR_WIDGET_SCRIPT_ID = "gdpr-consent-widget-script";
const GDPR_WIDGET_SCRIPT_URL = "/gpdr/gdpr-widget.iife.js";

function loadGdprWidget() {
  if (customElements.get("gdpr-consent-widget")) return Promise.resolve();
  return new Promise((resolve, reject) => {
    const existing = document.getElementById(GDPR_WIDGET_SCRIPT_ID);
    if (existing) {
      existing.addEventListener("load", resolve, { once: true });
      existing.addEventListener("error", reject, { once: true });
      return;
    }
    const script = document.createElement("script");
    script.id = GDPR_WIDGET_SCRIPT_ID;
    script.src = GDPR_WIDGET_SCRIPT_URL;
    script.async = true;
    script.addEventListener("load", resolve, { once: true });
    script.addEventListener("error", () => reject(new Error("Widget GDPR non disponibile")), { once: true });
    document.head.appendChild(script);
  });
}

export default function GdprWidget() {
  const [ready, setReady] = useState(() => Boolean(customElements.get("gdpr-consent-widget")));

  useEffect(() => {
    let active = true;
    loadGdprWidget()
      .then(() => active && setReady(true))
      .catch((error) => console.error(error));
    return () => { active = false; };
  }, []);

  return ready ? <gdpr-consent-widget /> : null;
}
