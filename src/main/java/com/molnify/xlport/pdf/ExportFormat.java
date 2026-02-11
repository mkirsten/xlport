package com.molnify.xlport.pdf;

/**
 * Class to handle export to different formats through Google. Used to export from Excel to PDF
 * Feature complete as of writing April 2019
 *
 * @author kirsten
 */
public class ExportFormat {
  // Flags taken from https://gist.github.com/Spencer-Easton/78f9867a691e549c9c70
  private int sheetId = -1;
  private int startRow = -1, endRow = -1, startColumn = -1, endColumn = -1;
  private float marginTop = -1, marginBottom = -1, marginLeft = -1, marginRight = -1;
  private String namedRange = null;
  private FormatEnum format = FormatEnum.PDF;
  private SizeEnum size = SizeEnum.A4;
  private AlignmentHorizontalEnum alignmentHorizontal = null;
  private AlignmentVerticalEnum alignmentVertical = null;
  private boolean portrait = true;
  private boolean printTitle = false;
  private boolean repeatRowHeaders = false;
  private boolean showPageNumbers = false;
  private boolean showGridLines = false;
  private boolean printNotes = false;
  private boolean attachment = true;

  public ExportFormat() {
    setFormat(FormatEnum.PDF);
  }

  public String getExportURLForId(String documentId) {
    StringBuffer urlString = new StringBuffer();
    urlString.append("https://docs.google.com/spreadsheets/d/");
    urlString.append(documentId);
    urlString.append("/export?");
    urlString.append("format=").append(format).append("&");

    // Some formats have no or limited options
    if (format == FormatEnum.XLSX || format == FormatEnum.ODF || format == FormatEnum.ZIP) {
      return urlString.toString();
    } else if (format == FormatEnum.CSV || format == FormatEnum.TSV) {
      if (sheetId < 0)
        throw new IllegalArgumentException("Need to set sheet when using CSV or TSV export");
      urlString.append("gid=").append(sheetId).append("&");
      return urlString.toString();
    }

    // Only PDF has more complex options
    urlString.append("size=").append(size).append("&");
    urlString.append("portrait=").append(portrait).append("&");
    urlString.append("printtitle=").append(printTitle).append("&");
    urlString.append("fzr=").append(repeatRowHeaders).append("&");
    urlString.append("printnotes=").append(printNotes).append("&");
    if (showPageNumbers) urlString.append("pagenum=").append("CENTER").append("&");
    if (showGridLines) urlString.append("gridlines=").append(showGridLines).append("&");
    if (sheetId >= 0) urlString.append("gid=").append(sheetId).append("&");

    // Only export a range, either named or with row/column identification
    if (namedRange != null) {
      if (sheetId < 0)
        throw new IllegalArgumentException("Need to set sheet when using export range");
      urlString.append("range=").append(namedRange).append("&");
    } else if (startRow > 0) {
      if (sheetId < 0)
        throw new IllegalArgumentException("Need to set sheet when using export range");
      urlString.append("ir=").append(false).append("&");
      urlString.append("ic=").append(false).append("&");
      urlString.append("r1=").append(startRow).append("&");
      urlString.append("c1=").append(endRow).append("&");
      urlString.append("r2=").append(startColumn).append("&");
      urlString.append("c2=").append(endColumn).append("&");
    }

    // All margins need to be set for this to work
    if (marginTop >= 0) {
      urlString.append("top_margin=").append(marginTop).append("&");
      urlString.append("bottom_margin=").append(marginBottom).append("&");
      urlString.append("left_margin=").append(marginLeft).append("&");
      urlString.append("right_margin=").append(marginRight).append("&");
    }
    if (alignmentHorizontal != null)
      urlString.append("horizontal_alignment=").append(alignmentHorizontal).append("&");
    if (alignmentVertical != null)
      urlString.append("vertical_alignment=").append(alignmentVertical).append("&");
    urlString.append("attachment=").append(attachment);
    return urlString.toString();
  }

  public ExportFormat setAttachment(boolean attachment) {
    this.attachment = attachment;
    return this;
  }

  public ExportFormat setAlignmentHorizontal(AlignmentHorizontalEnum h) {
    this.alignmentHorizontal = h;
    return this;
  }

  public ExportFormat setAlignmentVertical(AlignmentVerticalEnum v) {
    this.alignmentVertical = v;
    return this;
  }

  public ExportFormat setExportRange(String namedRange) {
    startRow = -1;
    endRow = -1;
    startColumn = -1;
    endColumn = -1;
    this.namedRange = namedRange;
    return this;
  }

  public ExportFormat setMargins(float top, float bottom, float left, float right) {
    if (top < 0 || bottom < 0 || left < 0 || right < 0)
      throw new IllegalArgumentException("Margins >0 required");
    this.marginTop = top;
    this.marginBottom = bottom;
    this.marginLeft = left;
    this.marginRight = right;
    return this;
  }

  public ExportFormat setExportRange(int startRow, int endRow, int startColumn, int endColumn) {
    if (startRow <= endRow
        || startColumn <= endColumn
        || startRow * endRow * startColumn * endColumn == 0)
      throw new IllegalArgumentException("Illegal export range");
    namedRange = null;
    this.startRow = startRow;
    this.endRow = endRow;
    this.startColumn = startColumn;
    this.endColumn = endColumn;
    return this;
  }

  public ExportFormat setPrintNotes(boolean printNotes) {
    this.printNotes = printNotes;
    return this;
  }

  public ExportFormat setShowGridLines(boolean showGridLines) {
    this.showGridLines = showGridLines;
    return this;
  }

  public ExportFormat setShowPageNumbers(boolean showPageNumbers) {
    this.showPageNumbers = showPageNumbers;
    return this;
  }

  public ExportFormat setPrintTitle(boolean printTitle) {
    this.printTitle = printTitle;
    return this;
  }

  public ExportFormat setPortrait(boolean portrait) {
    this.portrait = portrait;
    return this;
  }

  public ExportFormat setSize(SizeEnum size) {
    this.size = size;
    return this;
  }

  public ExportFormat setSheetId(int id) {
    this.sheetId = id;
    return this;
  }

  public ExportFormat setRepeatRowHeaders(boolean repeatRowHeaders) {
    this.repeatRowHeaders = repeatRowHeaders;
    return this;
  }

  public ExportFormat setFormat(FormatEnum format) {
    this.format = format;
    return this;
  }

  public enum AlignmentHorizontalEnum {
    LEFT,
    CENTER,
    RIGHT;
  }

  public enum AlignmentVerticalEnum {
    TOP,
    MIDDLE,
    BOTTOM;
  }

  public enum FormatEnum {
    PDF,
    XLSX,
    CSV,
    TSV,
    ODF,
    ZIP;

    public String toString() {
      return name().toLowerCase();
    }
  }

  public enum SizeEnum {
    LETTER,
    TABLOID,
    LEGAL,
    STATEMENT,
    EXECUTIVE,
    FOLIO,
    A3,
    A4,
    A5,
    B4,
    B5;

    public String toString() {
      return Integer.toString(ordinal());
    }
  }
}
