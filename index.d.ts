declare namespace XLSX {
  enum BorderStyle {
    /** No border */
    NONE_BORDER = 0,
    /** Thin border style */
    THIN_BORDER = 1,
    /** Medium border style */
    MEDIUM_BORDER = 2,
    /** Dashed border style */
    DASHED_BORDER = 3,
    /** Dotted border style */
    DOTTED_BORDER = 4,
    /** Thick border style */
    THICK_BORDER = 5,
    /** Double border style */
    DOUBLE_BORDER = 6,
    /** Hair border style */
    HAIR_BORDER = 7,
    /** Medium dashed border style */
    MEDIUM_DASHED_BORDER = 8,
    /** Dash-dot border style */
    DASH_DOT_BORDER = 9,
    /** Medium dash-dot border style */
    MEDIUM_DASH_DOT_BORDER = 10,
    /** Dash-dot-dot border style */
    DASH_DOT_DOT_BORDER = 11,
    /** Medium dash-dot-dot border style */
    MEDIUM_DASH_DOT_DOT_BORDER = 12,
    /** Slant dash-dot border style */
    SLANT_DASH_DOT_BORDER = 13,
  }

  enum ChartType {
    /** Area chart. */
    AREA_CHART = 1,
    /** Area chart - stacked. */
    AREA_STACKED_CHART = 2,
    /** Area chart - percentage stacked. */
    AREA_STACKED_PERCENT_CHART = 3,
    /** Bar chart. */
    BAR_CHART = 4,
    /** Bar chart - stacked. */
    BAR_STACKED_CHART = 5,
    /** Bar chart - percentage stacked. */
    BAR_STACKED_PERCENT_CHART = 6,
    /** Column chart. */
    COLUMN_CHART = 7,
    /** Column chart - stacked. */
    COLUMN_STACKED_CHART = 8,
    /** Column chart - percentage stacked. */
    COLUMN_STACKED_PERCENT_CHART = 9,
    /** Doughnut chart. */
    DOUGHNUT_CHART = 10,
    /** Line chart. */
    LINE_CHART = 11,
    /** Line chart - stacked. */
    LINE_STACKED_CHART = 12,
    /** Line chart - percentage stacked. */
    LINE_STACKED_PERCENT_CHART = 13,
    /** Pie chart. */
    PIE_CHART = 14,
    /** Scatter chart. */
    SCATTER_CHART = 15,
    /** Scatter chart - straight. */
    SCATTER_STRAIGHT_CHART = 16,
    /** Scatter chart - straight with markers. */
    SCATTER_STRAIGHT_WITH_MARKERS_CHART = 17,
    /** Scatter chart - smooth. */
    SCATTER_SMOOTH_CHART = 18,
    /** Scatter chart - smooth with markers. */
    SCATTER_SMOOTH_WITH_MARKERS_CHART = 19,
    /** Radar chart. */
    RADAR_CHART = 20,
    /** Radar chart - with markers. */
    RADAR_WITH_MARKERS_CHART = 21,
    /** Radar chart - filled. */
    RADAR_FILLED_CHART = 22,
  }

  class Worksheet {
    insertChart(row: number, column: number, chart: Chart): void;
    insertImage(row: number, column: number, image: Uint8Array): void;
    setColumn(
      firstColumn: number,
      lastColumn: number,
      width: number,
      format?: Format
    ): void;
    setRow(row: number, height: number, format?: Format): void;
    setFooter(footer: string): void;
    setHeader(header: string): void;
    writeDatetime(
      row: number,
      column: number,
      date: Date,
      format?: Format
    ): void;
    writeNumber(
      row: number,
      column: number,
      number: number,
      format?: Format
    ): void;
    writeString(
      row: number,
      column: number,
      string: string,
      format?: Format
    ): void;
  }

  type Chart = ExcelWriter.Chart;
  type Format = ExcelWriter.Format;
}

declare namespace ExcelWriter {
  class Chart {
    static AREA_CHART: XLSX.ChartType.AREA_CHART;
    static AREA_STACKED_CHART: XLSX.ChartType.AREA_STACKED_CHART;
    static AREA_STACKED_PERCENT_CHART: XLSX.ChartType.AREA_STACKED_PERCENT_CHART;
    static BAR_CHART: XLSX.ChartType.BAR_CHART;
    static BAR_STACKED_CHART: XLSX.ChartType.BAR_STACKED_CHART;
    static BAR_STACKED_PERCENT_CHART: XLSX.ChartType.BAR_STACKED_PERCENT_CHART;
    static COLUMN_CHART: XLSX.ChartType.COLUMN_CHART;
    static COLUMN_STACKED_CHART: XLSX.ChartType.COLUMN_STACKED_CHART;
    static COLUMN_STACKED_PERCENT_CHART: XLSX.ChartType.COLUMN_STACKED_PERCENT_CHART;
    static DOUGHNUT_CHART: XLSX.ChartType.DOUGHNUT_CHART;
    static LINE_CHART: XLSX.ChartType.LINE_CHART;
    static LINE_STACKED_CHART: XLSX.ChartType.LINE_STACKED_CHART;
    static LINE_STACKED_PERCENT_CHART: XLSX.ChartType.LINE_STACKED_PERCENT_CHART;
    static PIE_CHART: XLSX.ChartType.PIE_CHART;
    static SCATTER_CHART: XLSX.ChartType.SCATTER_CHART;
    static SCATTER_STRAIGHT_CHART: XLSX.ChartType.SCATTER_STRAIGHT_CHART;
    static SCATTER_STRAIGHT_WITH_MARKERS_CHART: XLSX.ChartType.SCATTER_STRAIGHT_WITH_MARKERS_CHART;
    static SCATTER_SMOOTH_CHART: XLSX.ChartType.SCATTER_SMOOTH_CHART;
    static SCATTER_SMOOTH_WITH_MARKERS_CHART: XLSX.ChartType.SCATTER_SMOOTH_WITH_MARKERS_CHART;
    static RADAR_CHART: XLSX.ChartType.RADAR_CHART;
    static RADAR_WITH_MARKERS_CHART: XLSX.ChartType.RADAR_WITH_MARKERS_CHART;
    static RADAR_FILLED_CHART: XLSX.ChartType.RADAR_FILLED_CHART;
    addSeries(categories: string | null, values: string): void;
    setTitleName(name: string): void;
    setTitleNameFont(font: ChartFont): void;
  }

  class Format {
    static NONE_BORDER: XLSX.BorderStyle.NONE_BORDER;
    static THIN_BORDER: XLSX.BorderStyle.THIN_BORDER;
    static MEDIUM_BORDER: XLSX.BorderStyle.MEDIUM_BORDER;
    static DASHED_BORDER: XLSX.BorderStyle.DASHED_BORDER;
    static DOTTED_BORDER: XLSX.BorderStyle.DOTTED_BORDER;
    static THICK_BORDER: XLSX.BorderStyle.THICK_BORDER;
    static DOUBLE_BORDER: XLSX.BorderStyle.DOUBLE_BORDER;
    static HAIR_BORDER: XLSX.BorderStyle.HAIR_BORDER;
    static MEDIUM_DASHED_BORDER: XLSX.BorderStyle.MEDIUM_DASHED_BORDER;
    static DASH_DOT_BORDER: XLSX.BorderStyle.DASH_DOT_BORDER;
    static MEDIUM_DASH_DOT_BORDER: XLSX.BorderStyle.MEDIUM_DASH_DOT_BORDER;
    static DASH_DOT_DOT_BORDER: XLSX.BorderStyle.DASH_DOT_DOT_BORDER;
    static MEDIUM_DASH_DOT_DOT_BORDER: XLSX.BorderStyle.MEDIUM_DASH_DOT_DOT_BORDER;
    static SLANT_DASH_DOT_BORDER: XLSX.BorderStyle.SLANT_DASH_DOT_BORDER;
    setBgColor(color: number): void;
    setFgColor(color: number): void;
    setBorderColor(color: number): void;
    setFontColor(color: number): void;
    setBold(): void;
    setBorder(style: BorderStyle): void;
    setNumFormat(format: string): void;
  }

  enum Color {
    /** Black */
    BLACK_COLOR = 0x1000000,
    /** Blue */
    BLUE_COLOR = 0x0000ff,
    /** Brown */
    BROWN_COLOR = 0x800000,
    /** Cyan */
    CYAN_COLOR = 0x00ffff,
    /** Gray */
    GRAY_COLOR = 0x808080,
    /** Green */
    GREEN_COLOR = 0x008000,
    /** Lime */
    LIME_COLOR = 0x00ff00,
    /** Magenta */
    MAGENTA_COLOR = 0xff00ff,
    /** Navy */
    NAVY_COLOR = 0x000080,
    /** Orange */
    ORANGE_COLOR = 0xff6600,
    /** Pink */
    PINK_COLOR = 0xff00ff,
    /** Purple */
    PURPLE_COLOR = 0x800080,
    /** Red */
    RED_COLOR = 0xff0000,
    /** Silver */
    SILVER_COLOR = 0xc0c0c0,
    /** White */
    WHITE_COLOR = 0xffffff,
    /** Yellow */
    YELLOW_COLOR = 0xffff00,
  }

  interface ChartFont {
    /** The chart font name, such as "Arial" or "Calibri". */
    name?: string;
    /** The chart font size. The default is 11. */
    size?: number;
    /** The chart font bold property. */
    bold?: boolean;
    /** The chart font italic property. */
    italic?: boolean;
    /** The chart font underline property. */
    underline?: boolean;
    /** The chart font rotation property. Range: -90 to 90, and 270, 271 and 360:
     *
     *  - The angles -90 to 90 are the normal range shown in the Excel user interface.
     *  - The angle 270 gives a stacked (top to bottom) alignment.
     *  - The angle 271 gives a stacked alignment for East Asian fonts.
     *  - The angle 360 gives an explicit angle of 0 to override the y axis default.
     * */
    rotation?: number;
    /** The chart font color. */
    color?: number;
    /** The chart font pitch family property. Rarely required, set to 0. */
    pitchFamily?: number;
    /** The chart font character set property. Rarely required, set to 0. */
    charset?: number;
    /** The chart font baseline property. Rarely required, set to 0. */
    baseline?: number;
  }

  class Workbook {
    addChart(type: ChartType): Chart;
    addFormat(): Format;
    addWorksheet(name: string): Worksheet;
    close(): ArrayBuffer;
  }

  type BorderStyle = XLSX.BorderStyle;
  type ChartType = XLSX.ChartType;
  type Worksheet = XLSX.Worksheet;
}

export = ExcelWriter;
