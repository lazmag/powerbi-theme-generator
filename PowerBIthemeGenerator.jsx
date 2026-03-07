import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  AlertCircle,
  CheckCircle2,
  Copy,
  Download,
  ImagePlus,
  MinusCircle,
  Monitor,
  Moon,
  Sparkles,
  Sun,
  Upload,
} from "lucide-react";

// ── Types ────────────────────────────────────────────────────────────────────

const themePresets = {
  lite: {
    id: "lite",
    name: "Lite",
    description: "Soft light theme with calm contrast and muted accents.",
    icon: Sun,
    background: "#F7F8FA",
    foreground: "#1F2937",
    panel: "#FFFFFF",
    border: "#E5E7EB",
    neutralText: "#8C8F96",
    paletteBias: 0.18,
    saturation: 0.82,
  },
  dusk: {
    id: "dusk",
    name: "Dusk",
    description: "Balanced dark theme designed for focused dashboard viewing.",
    icon: Moon,
    background: "#111827",
    foreground: "#E5E7EB",
    panel: "#17202F",
    border: "#2C3748",
    neutralText: "#A1A1AA",
    paletteBias: 0.28,
    saturation: 0.88,
  },
  mist: {
    id: "mist",
    name: "Mist",
    description: "Neutral grey-blue theme with understated, professional styling.",
    icon: Monitor,
    background: "#EEF2F5",
    foreground: "#243142",
    panel: "#FFFFFF",
    border: "#D8E0E8",
    neutralText: "#7E8794",
    paletteBias: 0.22,
    saturation: 0.8,
  },
  nightfall: {
    id: "nightfall",
    name: "Nightfall",
    description: "Muted charcoal theme with soft highlights and restrained colour use.",
    icon: Sparkles,
    background: "#0F172A",
    foreground: "#E2E8F0",
    panel: "#172033",
    border: "#2B3548",
    neutralText: "#9AA3AF",
    paletteBias: 0.3,
    saturation: 0.86,
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
  const background = hexToRgb(preset.background);
  const baseColours = ensureDistinct(safeHexes.map(hexToRgb), 52);

  const variants = [
    { index: 0, blendAmount: preset.paletteBias * 0.55, lightness: 1.0 },
    { index: 1, blendAmount: preset.paletteBias * 0.64, lightness: 1.04 },
    { index: 2, blendAmount: preset.paletteBias * 0.72, lightness: 0.98 },
    { index: 0, blendAmount: preset.paletteBias * 0.9, lightness: 1.12 },
    { index: 1, blendAmount: preset.paletteBias * 1.02, lightness: 0.9 },
    { index: 2, blendAmount: preset.paletteBias * 1.12, lightness: 1.08 },
    { index: 0, blendAmount: preset.paletteBias * 1.2, lightness: 0.84 },
    { index: 1, blendAmount: preset.paletteBias * 1.28, lightness: 1.16 },
  ];

  const paletteRgb = ensureDistinct(
    variants.map((v) => {
      const base = baseColours[v.index % baseColours.length];
      const current = rgbToHsl(base);
      const adjusted = withHsl(base, { s: current.s * preset.saturation, l: current.l * v.lightness });
      return blend(adjusted, background, Math.min(0.45, v.blendAmount));
    }),
    30,
  );

  const tintRef = baseColours.length > 1 ? blend(baseColours[0], baseColours[1], 0.35) : baseColours[0];
  const good    = blend(withHsl([92, 138, 103], { s: 0.34, l: 0.45 }), tintRef, 0.12);
  const neutral = blend(withHsl([166, 135, 88], { s: 0.34, l: 0.49 }), tintRef, 0.1);
  const bad     = blend(withHsl([171, 99, 101], { s: 0.38, l: 0.48 }), tintRef, 0.08);

  return {
    palette: paletteRgb.map(rgbToHex),
    good: rgbToHex(good),
    neutral: rgbToHex(neutral),
    bad: rgbToHex(bad),
  };
}

function buildThemeJson({ themeName, preset, fontFamily, baseText, titleText, smallText, showLegendTitle, roundedCorners, generatedColours }) {
  return {
    name: themeName,
    foreground: preset.foreground,
    background: preset.background,
    tableAccent: generatedColours.palette[0],
    dataColors: generatedColours.palette,
    good: generatedColours.good,
    neutral: generatedColours.neutral,
    bad: generatedColours.bad,
    visualStyles: {
      "*": {
        "*": {
          title: [{ show: true, fontFamily, fontSize: titleText, color: { solid: { color: preset.foreground } } }],
          background: [{ transparency: 100 }],
          visualHeader: [{ show: false }],
          border: [{ show: false, radius: roundedCorners }],
          labels: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily }],
        },
      },
      categoryAxis: {
        "*": {
          title: [{ show: false }],
          labels: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily }],
          gridlines: [{ show: false }],
        },
      },
      valueAxis: {
        "*": {
          title: [{ show: false }],
          labels: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily }],
          gridlines: [{ show: true, color: { solid: { color: preset.border } }, transparency: 65 }],
        },
      },
      legend: {
        "*": {
          title: [{ show: showLegendTitle }],
          labels: [{ color: { solid: { color: preset.foreground } }, fontSize: smallText, fontFamily }],
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
          grid: [{ gridVertical: true, gridHorizontal: false, outlineColor: { solid: { color: preset.border } }, fontSize: baseText, fontFamily }],
          values: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily }],
          columnHeaders: [{ color: { solid: { color: preset.foreground } }, fontSize: baseText, fontFamily, outline: "BottomOnly", outlineColor: { solid: { color: preset.border } } }],
        },
      },
      slicer: {
        "*": {
          title: [{ show: true, fontSize: baseText, fontFamily, color: { solid: { color: preset.foreground } } }],
          items: [{ fontSize: baseText, fontFamily, color: { solid: { color: preset.foreground } } }],
        },
      },
    },
    textClasses: {
      title:   { fontFace: fontFamily, fontSize: titleText,     color: preset.foreground },
      header:  { fontFace: fontFamily, fontSize: baseText,      color: preset.foreground },
      label:   { fontFace: fontFamily, fontSize: baseText,      color: preset.foreground },
      callout: { fontFace: fontFamily, fontSize: titleText + 6, color: preset.foreground },
    },
    page: {
      background: [{ color: { solid: { color: preset.background } }, transparency: 0 }],
      wallpaper:  [{ color: { solid: { color: preset.background } }, transparency: 0 }],
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
        const distinct = ensureDistinct(ranked, 54).slice(0, 3).map(rgbToHex);
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

  return (
    <div className="min-h-screen w-full overflow-auto bg-slate-50 p-6 text-slate-900">
      <div className="mx-auto w-full max-w-6xl space-y-6">

        {/* Header */}
        <Card>
          <CardHeader>
            <CardTitle className="text-3xl">Power BI Theme Generator</CardTitle>
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

                  <div className="min-h-[220px]">
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
                <div className="flex flex-col gap-3 md:flex-row md:items-start md:justify-between">
                  <div>
                    <CardTitle>4. Export JSON</CardTitle>
                    <CardDescription>Copy the JSON or download it and import it into Power BI Desktop.</CardDescription>
                  </div>
                  <div className="flex gap-2 shrink-0">
                    <Button variant="outline" onClick={copyJson}>
                      {copied ? <CheckCircle2 className="h-4 w-4 text-green-600" /> : <Copy className="h-4 w-4" />}
                      {copied ? "Copied!" : "Copy"}
                    </Button>
                    <Button onClick={() => downloadJson(`${themeName.replace(/\s+/g, "-").toLowerCase()}.json`, formattedJson)}>
                      <Download className="h-4 w-4" />
                      Download
                    </Button>
                  </div>
                </div>
              </CardHeader>
              <CardContent>
                <pre className="max-h-[420px] overflow-auto rounded-2xl bg-slate-900 p-4 text-[11px] leading-relaxed text-slate-200 font-mono whitespace-pre-wrap">
                  {formattedJson}
                </pre>
                <p className="mt-3 text-xs text-slate-500">
                  In Power BI Desktop: <span className="font-medium">View → Themes → Browse for themes</span> — then select the downloaded <code className="rounded bg-slate-100 px-1 py-0.5">.json</code> file.
                </p>
              </CardContent>
            </Card>
          </div>
        </div>
      </div>
    </div>
  );
}
