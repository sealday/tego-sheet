declare module 'fontkit' {
  export interface FontkitPath {
    transform(m0: number, m1: number, m2: number, m3: number, m4: number, m5: number): FontkitPath;
    toSVG(): string;
  }

  export interface FontkitGlyph {
    readonly path: FontkitPath;
  }

  export interface FontkitGlyphPosition {
    readonly xAdvance: number;
    readonly yAdvance: number;
    readonly xOffset: number;
    readonly yOffset: number;
  }

  export interface FontkitGlyphRun {
    readonly glyphs: readonly FontkitGlyph[];
    readonly positions: readonly FontkitGlyphPosition[];
    readonly advanceWidth: number;
  }

  export interface FontkitFont {
    readonly unitsPerEm: number;
    hasGlyphForCodePoint(codePoint: number): boolean;
    layout(text: string): FontkitGlyphRun;
  }

  export function create(buffer: Uint8Array): FontkitFont;
}
