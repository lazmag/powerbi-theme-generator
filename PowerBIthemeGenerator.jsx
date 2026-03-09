import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  CheckCircle2,
  Copy,
  Download,
  ImagePlus,
  Layers,
  MinusCircle,
  Monitor,
  Moon,
  Sun,
  Upload,
} from "lucide-react";

// ── Types ────────────────────────────────────────────────────────────────────

const themePresets = {
  lite: {
    id: "lite",
    name: "Lite",
    description: "Clean light canvas — brand colours lead the palette, neutral background keeps visuals sharp.",
    icon: Sun,
    background: "#F7F8FA",
    foreground: "#1F2937",
    panel: "#FFFFFF",
    border: "#E5E7EB",
    neutralText: "#8C8F96",
    paletteBias: 0.18,
    saturation: 0.82,
  },
  dark: {
    id: "dark",
    name: "Dark",
    description: "Dark canvas built for low-light dashboards — brand colours remain vivid and legible.",
    icon: Moon,
    background: "#111827",
    foreground: "#E5E7EB",
    panel: "#17202F",
    border: "#2C3748",
    neutralText: "#A1A1AA",
    paletteBias: 0.28,
    saturation: 0.88,
  },
  canvas: {
    id: "canvas",
    name: "Brand Canvas",
    description: "White visual panels over a brand-coloured background — pair with the PowerPoint canvas template for a modern split layout.",
    icon: Layers,
    background: "#F0F4F8",
    foreground: "#1A2535",
    panel: "#FFFFFF",
    border: "#D1DCE8",
    neutralText: "#6B7A8D",
    paletteBias: 0.18,
    saturation: 0.90,
  },
  contrast: {
    id: "contrast",
    name: "Contrast",
    description: "High-contrast white theme — maximum legibility for data-heavy tables, accessibility requirements, and print output.",
    icon: Monitor,
    background: "#FFFFFF",
    foreground: "#0F172A",
    panel: "#F8FAFC",
    border: "#94A3B8",
    neutralText: "#475569",
    paletteBias: 0.15,
    saturation: 0.88,
  },
};

const fontFamilies = ["Segoe UI", "Aptos", "Arial", "Calibri", "Verdana", "Tahoma"];
const hexPattern = /^#?[0-9A-Fa-f]{6}$/;

// ── Colour utilities ─────────────────────────────────────────────────────────

function clamp(value, min = 0, max = 255) {
  return Math.max(min, Math.min(max, value));
}

function normaliseHex(value) {
  const trimmed = value.trim();
  if (!trimmed) return "";
  const withHash = trimmed.startsWith("#") ? trimmed : `#${trimmed}`;
  if (!hexPattern.test(withHash)) return "";
  return withHash.toUpperCase();
}

function hexToRgb(hex) {
  const value = normaliseHex(hex).replace("#", "");
  return [
    parseInt(value.slice(0, 2), 16),
    parseInt(value.slice(2, 4), 16),
    parseInt(value.slice(4, 6), 16),
  ];
}

function rgbToHex([r, g, b]) {
  return `#${[r, g, b]
    .map((value) => clamp(Math.round(value)).toString(16).padStart(2, "0"))
    .join("")}`.toUpperCase();
}

function rgbToHsl([r, g, b]) {
  const red = r / 255, green = g / 255, blue = b / 255;
  const max = Math.max(red, green, blue), min = Math.min(red, green, blue);
  let h = 0, s = 0;
  const l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case red: h = (green - blue) / d + (green < blue ? 6 : 0); break;
      case green: h = (blue - red) / d + 2; break;
      default: h = (red - green) / d + 4; break;
    }
    h /= 6;
  }
  return { h, s, l };
}

function hslToRgb(h, s, l) {
  if (s === 0) { const g = clamp(Math.round(l * 255)); return [g, g, g]; }
  const hue2rgb = (p, q, t) => {
    let temp = t;
    if (temp < 0) temp += 1; if (temp > 1) temp -= 1;
    if (temp < 1 / 6) return p + (q - p) * 6 * temp;
    if (temp < 1 / 2) return q;
    if (temp < 2 / 3) return p + (q - p) * (2 / 3 - temp) * 6;
    return p;
  };
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  return [
    clamp(Math.round(hue2rgb(p, q, h + 1 / 3) * 255)),
    clamp(Math.round(hue2rgb(p, q, h) * 255)),
    clamp(Math.round(hue2rgb(p, q, h - 1 / 3) * 255)),
  ];
}

function blend(a, b, weight) {
  return [
    clamp(Math.round(a[0] * (1 - weight) + b[0] * weight)),
    clamp(Math.round(a[1] * (1 - weight) + b[1] * weight)),
    clamp(Math.round(a[2] * (1 - weight) + b[2] * weight)),
  ];
}

function withHsl(rgb, adjustments) {
  const current = rgbToHsl(rgb);
  return hslToRgb(
    adjustments.h ?? current.h,
    clamp((adjustments.s ?? current.s) * 100, 0, 100) / 100,
    clamp((adjustments.l ?? current.l) * 100, 6, 92) / 100,
  );
}

function relativeLuminance([r, g, b]) {
  const lin = (c) => { const s = c / 255; return s <= 0.03928 ? s / 12.92 : ((s + 0.055) / 1.055) ** 2.4; };
  return 0.2126 * lin(r) + 0.7152 * lin(g) + 0.0722 * lin(b);
}

function colourDistance(a, b) {
  return Math.sqrt((a[0] - b[0]) ** 2 + (a[1] - b[1]) ** 2 + (a[2] - b[2]) ** 2);
}

function ensureDistinct(colours, minDistance = 44) {
  const output = [];
  colours.forEach((colour) => {
    let candidate = colour, attempts = 0;
    while (output.some((e) => colourDistance(e, candidate) < minDistance) && attempts < 8) {
      const current = rgbToHsl(candidate);
      const dir = attempts % 2 === 0 ? 1.08 : 0.92;
      candidate = withHsl(candidate, { l: current.l * dir, s: current.s * 0.96 });
      attempts++;
    }
    output.push(candidate);
  });
  return output;
}

function defaultSourceColours() {
  return ["#4E79A7", "#59A14F", "#B07AA1"];
}

function buildGeneratedColours(sourceHexes, preset) {
  const baseHexes = sourceHexes.filter(Boolean).slice(0, 3);
  const safeHexes = (baseHexes.length ? baseHexes : defaultSourceColours())
    .map(normaliseHex).filter(Boolean);

  // Determine theme type from background luminance (WCAG-based)
  const bgLum = relativeLuminance(hexToRgb(preset.background));
  const isDark = bgLum < 0.2;

  // Lightness target band — colours must be legible on the background
  // Light bg → darker colours (30–56%), Dark bg → lighter colours (48–72%)
  const lightMin = isDark ? 0.48 : 0.30;
  const lightMax = isDark ? 0.72 : 0.56;
  const lightMid = (lightMin + lightMax) / 2;

  // Saturation bounds
  const satCap = preset.saturation;
  const satMin = isDark ? 0.55 : 0.48;

  const sourceRgb = ensureDistinct(safeHexes.map(hexToRgb), 50);
  const sourceHsl = sourceRgb.map(rgbToHsl);

  // Slots 0-2: adapted source colours — brand colours appear directly in the palette.
  // Lightness is clamped into the legibility band; saturation blended to preserve brand character.
  const adaptedSource = sourceRgb.map((rgb) => {
    const hsl = rgbToHsl(rgb);
    const l = Math.max(lightMin, Math.min(lightMax, hsl.l));
    const s = Math.max(satMin, Math.min(satCap, hsl.s * 0.65 + satCap * 0.35));
    return hslToRgb(hsl.h, s, l);
  });

  // Slots 3-7: hue-shifted variants and near-complementary fills
  const n = sourceHsl.length;
  const extraHues = [
    (sourceHsl[0].h + 0.10) % 1,
    (sourceHsl[1 % n].h - 0.08 + 1) % 1,
    (sourceHsl[2 % n].h + 0.13) % 1,
    (sourceHsl[0].h + 0.50) % 1,
    (sourceHsl[1 % n].h + 0.44) % 1,
  ];
  const lOffsets = [-0.04, +0.06, -0.07, +0.04, -0.06];
  const extraRgb = extraHues.map((h, i) => {
    const srcSat = sourceHsl[i % n].s;
    const sat = Math.max(satMin, Math.min(satCap, srcSat * 0.45 + satCap * 0.55));
    const l = Math.max(lightMin, Math.min(lightMax, lightMid + lOffsets[i]));
    return hslToRgb(h, sat, l);
  });

  const paletteRgb = ensureDistinct([...adaptedSource, ...extraRgb], 30);

  // Colourblind-friendly status colours (Okabe-Ito inspired).
  // Teal-green + amber + vermillion are distinguishable for deuteranopes (red-green deficiency).
  const statusL = isDark ? 0.62 : 0.42;
  const good    = hslToRgb(0.463, 0.82, statusL); // teal-green ~167°
  const neutral = hslToRgb(0.108, 0.88, statusL); // amber ~39°
  const bad     = hslToRgb(0.038, 0.78, statusL); // vermillion ~14°

  return {
    palette: paletteRgb.map(rgbToHex),
    good: rgbToHex(good),
    neutral: rgbToHex(neutral),
    bad: rgbToHex(bad),
  };
}

function buildThemeJson({ themeName, preset, fontFamily, baseText, titleText, smallText, showLegendTitle, roundedCorners, generatedColours }) {
  const isDark = relativeLuminance(hexToRgb(preset.background)) < 0.2;
  // Applied filter card bg: lifted above the pane on dark themes, depressed below panel on light themes
  const appliedCardBg = isDark ? preset.panel : preset.border;

  return {
    name: themeName,
    dataColors: generatedColours.palette,
    background: preset.background,
    foreground: preset.foreground,
    tableAccent: generatedColours.palette[0],
    good: generatedColours.good,
    neutral: generatedColours.neutral,
    bad: generatedColours.bad,
    // Full structural color set — required for dark themes to propagate correctly
    firstLevelElements: preset.foreground,
    secondLevelElements: preset.neutralText,
    thirdLevelElements: preset.border,
    fourthLevelElements: preset.neutralText,
    secondaryBackground: preset.panel,
    textClasses: {
      title:   { fontFace: fontFamily, fontSize: titleText,     color: preset.foreground },
      header:  { fontFace: fontFamily, fontSize: baseText,      color: preset.foreground },
      label:   { fontFace: fontFamily, fontSize: baseText,      color: preset.foreground },
      callout: { fontFace: fontFamily, fontSize: titleText + 6, color: preset.foreground },
    },
    visualStyles: {
      "*": {
        "*": {
          title: [{ show: true, fontSize: titleText, fontFamily, fontColor: { solid: { color: preset.foreground } } }],
          background: [{ transparency: 100 }],
          visualHeader: [{ show: false }],
          border: [{ show: false, cornerRadius: roundedCorners }],
          categoryAxis: [{
            showAxisTitle: false,
            labelColor: { solid: { color: preset.foreground } },
            fontSize: baseText,
            fontFamily,
            gridlineShow: false,
          }],
          valueAxis: [{
            show: false,
            showAxisTitle: false,
            gridlineShow: false,
          }],
          dataLabels: [{
            show: true,
            color: { solid: { color: preset.foreground } },
            fontSize: smallText,
            fontFamily,
            enableBackground: true,
            backgroundColor: { solid: { color: preset.panel } },
            backgroundTransparency: 20,
          }],
          legend: [{
            show: true,
            showTitle: showLegendTitle,
            labelColor: { solid: { color: preset.foreground } },
            fontSize: smallText,
            fontFamily,
          }],
          // Filter pane panel
          outspacePane: [{
            backgroundColor: { solid: { color: preset.panel } },
            foregroundColor: { solid: { color: preset.foreground } },
            borderColor: { solid: { color: preset.border } },
            transparency: 0,
            fontFamily,
            border: false,
          }],
          // Filter cards — Applied (active) clearly distinct, Available (inactive) subtle
          filterCard: [
            {
              $id: "Applied",
              backgroundColor: { solid: { color: appliedCardBg } },
              foregroundColor: { solid: { color: preset.foreground } },
              borderColor: { solid: { color: generatedColours.palette[0] } },
              inputBoxColor: { solid: { color: preset.panel } },
              transparency: 0,
              fontFamily,
              border: true,
              textSize: smallText,
            },
            {
              $id: "Available",
              backgroundColor: { solid: { color: preset.background } },
              foregroundColor: { solid: { color: preset.neutralText } },
              borderColor: { solid: { color: preset.border } },
              inputBoxColor: { solid: { color: preset.panel } },
              transparency: 0,
              fontFamily,
              border: false,
              textSize: smallText,
            },
          ],
        },
      },
      // Page canvas and outer area background
      page: {
        "*": {
          background: [{ color: { solid: { color: preset.background } }, transparency: 0 }],
          outspace: [{ color: { solid: { color: preset.panel } }, transparency: 0 }],
        },
      },
      card: {
        "*": {
          labels: [{ color: { solid: { color: preset.foreground } }, fontSize: titleText + 6, fontFamily }],
          categoryLabels: [{ color: { solid: { color: preset.neutralText } }, fontSize: smallText, fontFamily }],
        },
      },
      tableEx: {
        "*": {
          grid: [{ gridVertical: true, gridHorizontal: false, outlineColor: { solid: { color: preset.border } } }],
          values: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily }],
          columnHeaders: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily, outline: "BottomOnly", outlineColor: { solid: { color: preset.border } } }],
        },
      },
      slicer: {
        "*": {
          title: [{ show: true, fontSize: baseText, fontFamily, fontColor: { solid: { color: preset.foreground } } }],
          items: [{ fontSize: baseText, fontFamily, fontColor: { solid: { color: preset.foreground } } }],
        },
      },
    },
  };
}

function downloadJson(filename, content) {
  const blob = new Blob([content], { type: "application/json" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url; a.download = filename; a.click();
  URL.revokeObjectURL(url);
}

async function extractMainColoursFromFile(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onerror = () => reject(new Error("Unable to read file."));
    reader.onload = () => {
      const result = reader.result;
      if (typeof result !== "string") { reject(new Error("Unsupported file data.")); return; }
      const image = new Image();
      image.onload = () => {
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d", { willReadFrequently: true });
        if (!context) { reject(new Error("Unable to analyse the uploaded file.")); return; }
        const maxSize = 140;
        const scale = Math.min(maxSize / image.width, maxSize / image.height, 1);
        canvas.width = Math.max(1, Math.round(image.width * scale));
        canvas.height = Math.max(1, Math.round(image.height * scale));
        context.drawImage(image, 0, 0, canvas.width, canvas.height);
        const { data } = context.getImageData(0, 0, canvas.width, canvas.height);
        const buckets = new Map();
        for (let i = 0; i < data.length; i += 4) {
          const alpha = data[i + 3]; if (alpha < 120) continue;
          const rgb = [data[i], data[i + 1], data[i + 2]];
          const avg = (rgb[0] + rgb[1] + rgb[2]) / 3; if (avg < 18 || avg > 242) continue;
          const keyRgb = [Math.round(rgb[0]/24)*24, Math.round(rgb[1]/24)*24, Math.round(rgb[2]/24)*24];
          const key = keyRgb.join("-");
          const cur = buckets.get(key);
          if (cur) cur.count++; else buckets.set(key, { count: 1, rgb: keyRgb });
        }
        const ranked = Array.from(buckets.values()).sort((a, b) => b.count - a.count).map(e => e.rgb);
        const distinct = ensureDistinct(ranked, 54).slice(0, 3).map(rgbToHex)
          .sort((a, b) => relativeLuminance(hexToRgb(a)) - relativeLuminance(hexToRgb(b)));
        resolve(distinct.length ? distinct : defaultSourceColours());
      };
      image.onerror = () => reject(new Error("Unable to load the uploaded image."));
      image.src = result;
    };
    reader.readAsDataURL(file);
  });
}

// ── Tiny UI primitives (no external dep) ─────────────────────────────────────

const Card = ({ className = "", children, style }) => (
  <div className={`rounded-3xl border-0 shadow-sm bg-white ${className}`} style={style}>{children}</div>
);
const CardHeader = ({ children }) => <div className="px-6 pt-6 pb-2">{children}</div>;
const CardTitle = ({ children, className = "" }) => <h2 className={`text-xl font-semibold text-slate-900 ${className}`}>{children}</h2>;
const CardDescription = ({ children }) => <p className="text-sm text-slate-500 mt-1">{children}</p>;
const CardContent = ({ children }) => <div className="px-6 pb-6">{children}</div>;

const Label = ({ children, htmlFor, className = "" }) => (
  <label htmlFor={htmlFor} className={`text-sm font-medium text-slate-700 ${className}`}>{children}</label>
);

const Input = ({ value, onChange, placeholder, className = "" }) => (
  <input
    value={value}
    onChange={onChange}
    placeholder={placeholder}
    className={`w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-slate-300 bg-white ${className}`}
  />
);

const Button = ({ children, onClick, variant = "default", className = "", type = "button", disabled }) => {
  const base = "inline-flex items-center gap-2 rounded-2xl px-4 py-2 text-sm font-medium transition";
  const styles = variant === "outline"
    ? "border border-slate-200 bg-white text-slate-700 hover:bg-slate-50"
    : "bg-slate-900 text-white hover:bg-slate-700";
  return (
    <button type={type} onClick={onClick} disabled={disabled} className={`${base} ${styles} ${className}`}>
      {children}
    </button>
  );
};

const Switch = ({ checked, onCheckedChange }) => (
  <button
    role="switch"
    aria-checked={checked}
    onClick={() => onCheckedChange(!checked)}
    className={`relative inline-flex h-6 w-11 items-center rounded-full transition-colors ${checked ? "bg-slate-800" : "bg-slate-300"}`}
  >
    <span className={`inline-block h-4 w-4 rounded-full bg-white shadow transition-transform ${checked ? "translate-x-6" : "translate-x-1"}`} />
  </button>
);

const Slider = ({ min, max, step, value, onValueChange }) => (
  <input
    type="range" min={min} max={max} step={step} value={value[0]}
    onChange={(e) => onValueChange([Number(e.target.value)])}
    className="w-full accent-slate-800"
  />
);

// ── Main component ────────────────────────────────────────────────────────────

export default function PowerBIThemeGeneratorApp() {
  const [presetId, setPresetId] = useState("lite");
  const [themeName, setThemeName] = useState("Minimal Accessible Theme");
  const [fontFamily, setFontFamily] = useState("Segoe UI");
  const [baseText, setBaseText] = useState(11);
  const [titleText, setTitleText] = useState(14);
  const [smallText, setSmallText] = useState(10);
  const [showLegendTitle, setShowLegendTitle] = useState(false);
  const [roundedCorners, setRoundedCorners] = useState(6);
  const [copied, setCopied] = useState(false);
  const [pptxLoading, setPptxLoading] = useState(false);

  const [sourceMode, setSourceMode] = useState("logo");
  const [manualHexes, setManualHexes] = useState(["", "", ""]);
  const [logoFileName, setLogoFileName] = useState("");
  const [uploadedLogoColours, setUploadedLogoColours] = useState(defaultSourceColours());
  const [isDragActive, setIsDragActive] = useState(false);
  const [logoError, setLogoError] = useState("");

  const fileInputRef = useRef(null);
  const preset = themePresets[presetId];

  const activeSourceColours = useMemo(() => {
    if (sourceMode === "manual") {
      const validHexes = manualHexes.map(normaliseHex).filter(Boolean);
      return validHexes.length ? validHexes : defaultSourceColours();
    }
    return uploadedLogoColours.length ? uploadedLogoColours : defaultSourceColours();
  }, [manualHexes, sourceMode, uploadedLogoColours]);

  const generatedColours = useMemo(
    () => buildGeneratedColours(activeSourceColours, preset),
    [activeSourceColours, preset],
  );

  const themeJson = useMemo(
    () => buildThemeJson({ themeName, preset, fontFamily, baseText, titleText, smallText, showLegendTitle, roundedCorners, generatedColours }),
    [themeName, preset, fontFamily, baseText, titleText, smallText, showLegendTitle, roundedCorners, generatedColours],
  );

  const formattedJson = useMemo(() => JSON.stringify(themeJson, null, 2), [themeJson]);

  const copyJson = async () => {
    await navigator.clipboard.writeText(formattedJson);
    setCopied(true);
    window.setTimeout(() => setCopied(false), 1500);
  };

  const processLogoFile = async (file) => {
    const supportedTypes = ["image/png", "image/jpeg", "image/svg+xml"];
    const isSvg = file.name.toLowerCase().endsWith(".svg");
    if (!supportedTypes.includes(file.type) && !isSvg) {
      setLogoError("Please upload a PNG, JPG, JPEG or SVG logo file.");
      return;
    }
    try {
      setLogoError("");
      setLogoFileName(file.name);
      const colours = await extractMainColoursFromFile(file);
      setUploadedLogoColours(colours);
      setSourceMode("logo");
    } catch (error) {
      setLogoError(error instanceof Error ? error.message : "Unable to extract colours from the file.");
    }
  };

  useEffect(() => { if (sourceMode === "manual") setLogoError(""); }, [sourceMode]);

  const handleDownloadPptx = async () => {
    setPptxLoading(true);
    try {
      const { default: JSZip } = await import("jszip");

      // Fetch the appropriate template PPTX
      const base = import.meta.env.BASE_URL;
      const templateUrl = `${base}templates/${preset.id}.pptx`;
      const res = await fetch(templateUrl);
      if (!res.ok) throw new Error(`Failed to load template: ${res.status}`);
      const buf = await res.arrayBuffer();
      const zip = await JSZip.loadAsync(buf);

      // ── Compute colour substitution map ──────────────────────────────────────
      const hxUp = (c) => c.replace("#", "").toUpperCase();

      // Brand-derived dark bg shades (for canvas/contrast — all colours adapt)
      const p0hsl = rgbToHsl(hexToRgb(generatedColours.palette[0]));
      const toHex6up = (rgb) => rgbToHex(rgb).replace("#", "").toUpperCase();
      const BD1 = toHex6up(hslToRgb(p0hsl.h, Math.min(0.90, p0hsl.s * 0.70 + 0.20), 0.10));
      const BD2 = toHex6up(hslToRgb(p0hsl.h, Math.min(0.85, p0hsl.s * 0.60 + 0.15), 0.19));

      const P0 = hxUp(generatedColours.palette[0]);
      const P1 = hxUp(generatedColours.palette[1]);
      const P2 = hxUp(generatedColours.palette[2] || generatedColours.palette[0]);
      const FG = hxUp(preset.foreground);

      // Maps: source colour (from template) → replacement colour (from palette)
      // lite/dark: background stays, only accent colours replace
      // canvas/contrast: all colours (incl. background) replace
      const colourMaps = {
        lite: {
          "00A1DF": P0,  // active tab, accent divider bar
          "00A3E0": P0,  // theme accent1 — logo text (scheme:accent1)
          "1B365D": FG,  // title text
        },
        dark: {
          "3AD1C5": P0,  // primary teal (KPI panel, data refresh)
          "33BBE9": P2,  // blue (3rd overview column header)
          "F0B323": P1,  // amber (2nd overview column + theme accent6)
          "5F8E8E": P2,  // muted secondary teal text
          // 140A41, 123266 kept — background stays as designed
        },
        canvas: {
          "33BBE9": P0,  // teal accent line and nav card tops
          "140A41": BD1, // darkest bg shade
          "123266": BD2, // panel bg and gradient end
          "113468": BD2, // thin decorative bar
        },
        contrast: {
          "33BBE9": P0,  // teal accent line and nav card tops
          "140A41": BD1,
          "123266": BD2,
          "113468": BD2,
        },
      };
      const colourMap = colourMaps[preset.id] ?? colourMaps.lite;

      // ── Apply replacements to all XML files in the zip ────────────────────────
      const xmlNames = Object.keys(zip.files).filter(
        (name) => name.endsWith(".xml") && !zip.files[name].dir,
      );
      for (const name of xmlNames) {
        let xml = await zip.files[name].async("string");
        let updated = xml;
        for (const [from, to] of Object.entries(colourMap)) {
          updated = updated.replace(new RegExp(`val="${from}"`, "gi"), `val="${to}"`);
        }
        if (updated !== xml) zip.file(name, updated);
      }

      // ── Generate and trigger download ─────────────────────────────────────────
      const slug = themeName.replace(/\s+/g, "-").toLowerCase();
      const blob = await zip.generateAsync({
        type: "blob",
        compression: "DEFLATE",
        compressionOptions: { level: 6 },
      });
      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = `${slug}-background-template.pptx`;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

    } finally {
      setPptxLoading(false);
    }
  };

  return (
    <div className="min-h-screen w-full overflow-auto bg-slate-50 p-6 text-slate-900">
      <div className="mx-auto w-full max-w-6xl space-y-6">

        {/* Header */}
        <Card>
          <CardHeader>
            <CardTitle className="text-5xl font-bold">Power BI Theme Generator</CardTitle>
            <CardDescription>
              Create restrained, colour-blind-aware Power BI theme JSON files with cleaner defaults for report pages and visuals.
            </CardDescription>
          </CardHeader>
        </Card>

        <div className="grid gap-6 lg:grid-cols-[1.1fr_0.9fr]">
          {/* ── Left column ── */}
          <div className="space-y-6">

            {/* Step 1 – Brand colours */}
            <Card>
              <CardHeader>
                <CardTitle>1. Add brand colours</CardTitle>
                <CardDescription>Upload a logo to extract main colours or enter up to three hex codes manually.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="space-y-4">
                  <div className="flex gap-4">
                    {[["logo","Use logo file"],["manual","Provide hex codes"]].map(([val, label]) => (
                      <label key={val} className="flex items-center gap-2 text-sm font-medium cursor-pointer">
                        <input type="radio" name="source-mode" checked={sourceMode === val} onChange={() => setSourceMode(val)} className="h-4 w-4" />
                        {label}
                      </label>
                    ))}
                  </div>

                  <div>
                    {sourceMode === "logo" ? (
                      <>
                        <div
                          className={`rounded-3xl border border-dashed p-5 transition ${isDragActive ? "border-slate-400 bg-slate-100" : "border-slate-300 bg-slate-50"}`}
                          onDragOver={(e) => { e.preventDefault(); setIsDragActive(true); }}
                          onDragLeave={() => setIsDragActive(false)}
                          onDrop={(e) => { e.preventDefault(); setIsDragActive(false); const f = e.dataTransfer.files?.[0]; if (f) void processLogoFile(f); }}
                        >
                          <div className="flex flex-col items-start gap-4">
                            <div className="flex h-12 w-12 items-center justify-center rounded-2xl bg-white shadow-sm">
                              <ImagePlus className="h-5 w-5" />
                            </div>
                            <div className="space-y-1">
                              <p className="font-medium">Upload logo file</p>
                              <p className="text-sm text-slate-600">PNG, JPG, JPEG or SVG. You can also drag and drop the file here.</p>
                            </div>
                            <div className="flex flex-wrap items-center gap-3">
                              <Button variant="outline" onClick={() => fileInputRef.current?.click()}>
                                <Upload className="h-4 w-4" /> Choose file
                              </Button>
                              {logoFileName && <span className="text-sm text-slate-600">{logoFileName}</span>}
                            </div>
                            <input ref={fileInputRef} type="file" accept=".png,.jpg,.jpeg,.svg,image/png,image/jpeg,image/svg+xml" className="hidden"
                              onChange={(e) => { const f = e.target.files?.[0]; if (f) void processLogoFile(f); }} />
                          </div>
                        </div>
                        {logoError && <p className="mt-2 text-sm text-red-600">{logoError}</p>}

                        {/* Extracted colours preview */}
                        {uploadedLogoColours.length > 0 && logoFileName && (
                          <div className="mt-4 flex gap-3">
                            {uploadedLogoColours.map((c, i) => (
                              <div key={i} className="flex items-center gap-2">
                                <div className="h-8 w-8 rounded-xl border border-slate-200" style={{ backgroundColor: c }} />
                                <span className="text-xs font-mono text-slate-600">{c}</span>
                              </div>
                            ))}
                          </div>
                        )}
                      </>
                    ) : (
                      <div className="space-y-4">
                        <div className="grid gap-4 md:grid-cols-3">
                          {[0, 1, 2].map((index) => {
                            const previewColour = normaliseHex(manualHexes[index]);
                            return (
                              <div key={index} className="space-y-2">
                                <Label>Colour {index + 1}</Label>
                                <Input
                                  value={manualHexes[index]}
                                  onChange={(e) => {
                                    const next = [...manualHexes];
                                    next[index] = e.target.value;
                                    setManualHexes(next);
                                  }}
                                  placeholder="#4E79A7"
                                />
                                <div className="rounded-2xl border border-slate-200 bg-slate-50 p-3">
                                  <div
                                    className="h-12 rounded-xl border"
                                    style={{ backgroundColor: previewColour || "#FFFFFF", borderColor: "#E5E7EB" }}
                                  />
                                  {previewColour
                                    ? <p className="mt-1 text-center text-xs font-mono text-slate-500">{previewColour}</p>
                                    : <p className="mt-1 text-center text-xs text-slate-400">Enter a hex code</p>
                                  }
                                </div>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    )}
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Step 2 – Style */}
            <Card>
              <CardHeader>
                <CardTitle>2. Choose a style</CardTitle>
                <CardDescription>All presets are minimal, muted, and designed to avoid harsh contrast or overly vivid tones.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid gap-4 md:grid-cols-2">
                  {Object.values(themePresets).map((item) => {
                    const Icon = item.icon;
                    const isSelected = presetId === item.id;
                    return (
                      <label key={item.id} className={`flex cursor-pointer items-start gap-4 rounded-2xl border p-4 transition hover:shadow-sm ${isSelected ? "border-slate-800 bg-slate-50" : "border-slate-200 bg-white"}`}>
                        <input type="radio" name="preset" value={item.id} checked={isSelected} onChange={() => setPresetId(item.id)} className="mt-1 h-4 w-4" />
                        <div className="flex-1 space-y-1">
                          <div className="flex items-center gap-2">
                            <Icon className="h-4 w-4" />
                            <span className="font-medium">{item.name}</span>
                          </div>
                          <p className="text-sm text-slate-600">{item.description}</p>
                        </div>
                      </label>
                    );
                  })}
                </div>
              </CardContent>
            </Card>

            {/* Step 3 – Tune */}
            <Card>
              <CardHeader>
                <CardTitle>3. Tune the theme</CardTitle>
                <CardDescription>Set typography and a few simple formatting options.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="grid gap-6 md:grid-cols-2">
                  <div className="space-y-5">
                    <div className="space-y-2">
                      <Label>Theme name</Label>
                      <Input value={themeName} onChange={(e) => setThemeName(e.target.value)} />
                    </div>
                    <div className="space-y-2">
                      <div className="flex items-center justify-between text-sm">
                        <Label>Corner radius</Label>
                        <span className="text-slate-500">{roundedCorners}px</span>
                      </div>
                      <Slider min={0} max={16} step={1} value={[roundedCorners]} onValueChange={(v) => setRoundedCorners(v[0] ?? 6)} />
                    </div>
                    <div className="flex items-center justify-between rounded-2xl border border-slate-200 p-4">
                      <div>
                        <p className="font-medium">Hide legend titles</p>
                        <p className="text-sm text-slate-600">Removes extra label noise from chart legends.</p>
                      </div>
                      <Switch checked={!showLegendTitle} onCheckedChange={(checked) => setShowLegendTitle(!checked)} />
                    </div>
                  </div>
                  <div className="space-y-5">
                    <div className="space-y-2">
                      <Label>Font family</Label>
                      <select value={fontFamily} onChange={(e) => setFontFamily(e.target.value)}
                        className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm bg-white outline-none focus:ring-2 focus:ring-slate-300">
                        {fontFamilies.map((f) => <option key={f} value={f}>{f}</option>)}
                      </select>
                    </div>
                    {[["Title text size", titleText, setTitleText, 12, 20], ["Base text size", baseText, setBaseText, 9, 14], ["Small text size", smallText, setSmallText, 8, 12]].map(([label, val, setter, min, max]) => (
                      <div key={label} className="space-y-2">
                        <div className="flex items-center justify-between text-sm">
                          <Label>{label}</Label>
                          <span className="text-slate-500">{val}px</span>
                        </div>
                        <Slider min={min} max={max} step={1} value={[val]} onValueChange={(v) => setter(v[0])} />
                      </div>
                    ))}
                  </div>
                </div>
              </CardContent>
            </Card>
          </div>

          {/* ── Right column ── */}
          <div className="space-y-6">

            {/* Preview */}
            <Card>
              <CardHeader>
                <CardTitle>Preview</CardTitle>
                <CardDescription>Palette and styling direction for the selected theme.</CardDescription>
              </CardHeader>
              <CardContent>
                <div
                  className="overflow-hidden rounded-[28px] border"
                  style={{ background: `linear-gradient(180deg, ${preset.background} 0%, ${preset.panel} 100%)`, color: preset.foreground, borderColor: preset.border }}
                >
                  <div className="border-b p-5" style={{ borderColor: preset.border }}>
                    <p className="text-xs font-medium opacity-60 mb-1">Selected style</p>
                    <div className="flex items-center gap-2 text-sm font-semibold">
                      {React.createElement(preset.icon, { className: "h-4 w-4" })}
                      {preset.name}
                    </div>
                  </div>

                  <div className="p-5 space-y-5">
                    {/* Palette */}
                    <div>
                      <p className="text-xs font-medium opacity-60 mb-3">Generated colour palette</p>
                      <div className="grid grid-cols-4 gap-2">
                        {generatedColours.palette.map((colour, index) => (
                          <div key={`${colour}-${index}`} className="space-y-1">
                            <div className="h-10 rounded-xl border" style={{ backgroundColor: colour, borderColor: "rgba(255,255,255,0.08)" }} />
                            <p className="text-[10px] font-mono opacity-70 truncate">{colour}</p>
                          </div>
                        ))}
                      </div>
                    </div>

                    {/* Status colours */}
                    <div>
                      <p className="text-xs font-medium opacity-60 mb-3">Status colours</p>
                      <div className="space-y-2">
                        {[
                          { label: "Good",    value: generatedColours.good,    Icon: CheckCircle2 },
                          { label: "Neutral", value: generatedColours.neutral, Icon: MinusCircle },
                          { label: "Bad",     value: generatedColours.bad,     Icon: AlertCircle },
                        ].map(({ label, value, Icon }) => (
                          <div key={label} className="flex items-center gap-3">
                            <div className="flex h-9 w-9 items-center justify-center rounded-xl border" style={{ backgroundColor: value, borderColor: "rgba(255,255,255,0.08)" }}>
                              <Icon className="h-4 w-4 text-white" />
                            </div>
                            <div>
                              <p className="text-[11px] font-medium opacity-60">{label}</p>
                              <p className="text-xs font-mono font-semibold">{value}</p>
                            </div>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>
                </div>
              </CardContent>
            </Card>

            {/* Step 4 – Export */}
            <Card>
              <CardHeader>
                <CardTitle>4. Export</CardTitle>
                <CardDescription>Download the theme JSON for Power BI, or a PowerPoint background template to customise and use as a canvas image.</CardDescription>
              </CardHeader>
              <CardContent>
                <div className="flex flex-wrap gap-2 mb-4">
                  <Button variant="outline" onClick={copyJson}>
                    {copied ? <CheckCircle2 className="h-4 w-4 text-green-600" /> : <Copy className="h-4 w-4" />}
                    {copied ? "Copied!" : "Copy JSON"}
                  </Button>
                  <Button variant="outline" onClick={handleDownloadPptx} disabled={pptxLoading}>
                    <Download className="h-4 w-4" />
                    {pptxLoading ? "Generating…" : "PowerPoint"}
                  </Button>
                  <Button onClick={() => downloadJson(`${themeName.replace(/\s+/g, "-").toLowerCase()}.json`, formattedJson)}>
                    <Download className="h-4 w-4" />
                    Download JSON
                  </Button>
                </div>
                <div className="space-y-1 text-xs text-slate-500">
                  <p><span className="font-medium">Theme JSON:</span> in Power BI Desktop go to <span className="font-medium">View → Themes → Browse for themes</span> and select the <code className="rounded bg-slate-100 px-1 py-0.5">.json</code> file.</p>
                  <p><span className="font-medium">PowerPoint template:</span> open the <code className="rounded bg-slate-100 px-1 py-0.5">.pptx</code>, customise slides 2 &amp; 3, then right-click each slide → <span className="font-medium">Save as Picture → SVG</span> (preferred) or PNG. Apply in Power BI: Format pane → Canvas background → Image. Set canvas size to <span className="font-medium">1600 × 900 px</span>.</p>
                </div>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
}
