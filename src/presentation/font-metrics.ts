/** Device-independent metrics for one resolved font face. */
export interface FontFaceMetrics {
  /** Average glyph advance at a ten-unit font size. */
  readonly averageAdvance: number;
  /** Stable line height at a ten-unit font size. */
  readonly lineHeight: number;
}

/** Explicit deterministic font metric input. */
export interface FontMetricsOptions {
  /** Known font faces keyed by family name. */
  readonly fonts: Readonly<Record<string, FontFaceMetrics>>;
  /** Family name emitted when a requested font is unavailable. */
  readonly fallbackFont: string;
  /** Metrics used by the fallback font. */
  readonly fallback: FontFaceMetrics;
}

/** A resolved font face and its deterministic measurements. */
export interface ResolvedFontMetrics extends FontFaceMetrics {
  /** Resolved family used for layout and output. */
  readonly fontFamily: string;
  /** Whether resolution selected the configured fallback. */
  readonly missing: boolean;
}

/** Font metric service without browser font or Canvas dependencies. */
export interface FontMetrics {
  /** Resolves a requested family to explicit metrics. */
  resolve(fontFamily: string): ResolvedFontMetrics;
  /** Measures text without reading browser or Canvas font state. */
  measure(text: string, fontFamily: string, fontSize: number): number;
}

/** Creates deterministic font metrics with one explicit fallback. */
export function createFontMetrics(options: FontMetricsOptions): FontMetrics {
  const fonts = new Map(Object.entries(options.fonts));
  const resolve = (fontFamily: string): ResolvedFontMetrics => {
    const metrics = fonts.get(fontFamily);
    return Object.freeze(
      metrics === undefined
        ? { fontFamily: options.fallbackFont, ...options.fallback, missing: true }
        : { fontFamily, ...metrics, missing: false },
    );
  };
  const metrics: FontMetrics = {
    resolve,
    measure(text: string, fontFamily: string, fontSize: number) {
      const metrics = resolve(fontFamily);
      return text.length * metrics.averageAdvance * (fontSize / 10);
    },
  };
  return Object.freeze(metrics);
}
