package com.molnify.xlport.core;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.net.HttpURLConnection;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.ZonedDateTime;
import java.time.format.DateTimeFormatter;
import java.time.temporal.ChronoUnit;
import java.time.temporal.TemporalAccessor;
import java.util.*;
import java.util.logging.Logger;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.model.CalculationChain;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONObject;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidation;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTDataValidations;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTWorksheet;

/**
 * Utility methods for Excel workbook manipulation, date handling, cell operations,
 * and workbook comparison.
 */
public class Utils {

  private static final Logger log = Logger.getLogger(Utils.class.getName());

  /**
   * This moves a spec part in JSON, back to the root properties - put under the sheetName in the
   * root tables - put directly under root (with collapsed name)
   *
   * @param sheetSpec
   * @param root
   */
  public static void translateSheetTemplateSpecToRoot(JSONObject sheetSpec, JSONObject root) {
    String sheetName = sheetSpec.getString("name");
    JSONObject data = sheetSpec.getJSONObject("data");
    Set<String> keys = data.keySet();
    if (!root.has(sheetName)) root.put(sheetName, new JSONObject());
    JSONObject sheetScoped = root.getJSONObject(sheetName);
    sheetScoped.put("_xlport_metadata", "sheet");
    for (String key : keys) {
      // Plain data in a JSONArray (the table case)
      if (data.get(key) instanceof JSONArray) {
        root.put(Utils.collapseName(sheetName, key), data.get(key));
      }
      // Data as a JSONObject, which is a table with both a spec for e.g., columns as well as data
      else if (data.get(key) instanceof JSONObject
          && data.getJSONObject(key).has("data")
          && data.getJSONObject(key).get("data") instanceof JSONArray) {
        root.put(Utils.collapseName(sheetName, key), data.getJSONObject(key).getJSONArray("data"));
      }
      // Property case
      else {
        sheetScoped.put(key, data.get(key));
      }
    }
    log.fine("Sheet spec: " + sheetSpec);
    log.fine("Root: " + root);
  }

  // Used to create a new unique name for an item (named range or tablename)
  public static String collapseName(String sheetName, String itemName) {
    return "_"
        + sheetName
            .replace("_", "..us..")
            .replace(" ", "__")
            .replace("-", "MNS")
            .replace("+", "PLS")
        + "_"
        + itemName;
  }

  // Used to split out sheet and item name
  public static String[] expandName(String collapsedName) {
    String[] parts = collapsedName.split("(?<!_)_(?!_)");
    if (parts.length < 3) return new String[] {collapsedName};
    String sheetFragment =
        parts[1].replace("__", " ").replace("..us..", "_").replace("MNS", "-").replace("PLS", "+");
    String itemName = parts[2];
    return new String[] {sheetFragment, itemName};
  }

  public static boolean isFormattedAsDate(String potentialDate) {
    // Perform quick check first, most likely it is not a date
    if (potentialDate == null
        || potentialDate.length() < 23
        || '-' != potentialDate.charAt(4)
        || '-' != potentialDate.charAt(7)
        || 'T' != potentialDate.charAt(10)) return false;
    try {
      jakarta.xml.bind.DatatypeConverter.parseDateTime(potentialDate);
      return true;
    } catch (IllegalArgumentException e) {
      return false;
    }
  }

  /**
   * Surprisingly complex. This implementation is taken from
   * https://stackoverflow.com/questions/2201925/converting-iso-8601-compliant-string-to-java-util-date/60214805#60214805
   * Should be a ISO 8601 compliant cast from a date as String to a Java Date object. The messy
   * thing is the timezone which often gets 1 hour wrong
   *
   * @param date A ISO 8601 compliant string
   * @return A Java Date. Null if the string couldn't be parsed
   */
  public static Date getAsDate(String date) {
    int implementation = 1;

    if (implementation == 1) {
      try {
        return new SimpleDateFormat("yyyy-MM-dd'T'HH:mm:ss").parse(date);
      } catch (ParseException e) {
        return null;
      }
    } else if (implementation == 2) {
      try {
        TemporalAccessor ta = DateTimeFormatter.ISO_INSTANT.parse(date);
        Instant i = Instant.from(ta);
        return Date.from(i.minus(1, ChronoUnit.HOURS));
      } catch (Exception e) {
        return null;
      }
    } else {
      if (isFormattedAsDate(date)) {
        // This is a bit tricky. Think the withoutTimezone is returned in UTC. When we produce a
        // date from it, it will adjust for daylight
        // Hence, check if we are currently in daylight and if so adjust back one hour
        Instant withoutTimezone =
            jakarta.xml.bind.DatatypeConverter.parseDateTime(date).toInstant();
        boolean jvmInDaylight = TimeZone.getDefault().inDaylightTime(new Date());
        // Still not resolved - both at -1 adjustment which is weird, but tests pass now
        // (2022-09-08)
        if (jvmInDaylight) return Date.from(withoutTimezone.minus(1, ChronoUnit.HOURS));
        else return Date.from(withoutTimezone.minus(1, ChronoUnit.HOURS));
      }
    }
    return null;
  }

  public static String getAsISO8601String(Date date) {
    Calendar c = Calendar.getInstance();
    c.setTime(date);
    ZonedDateTime d =
        LocalDate.of(c.get(Calendar.YEAR), c.get(Calendar.MONTH) + 1, c.get(Calendar.DAY_OF_MONTH))
            .atTime(
                c.get(Calendar.HOUR),
                c.get(Calendar.MINUTE),
                c.get(Calendar.SECOND),
                c.get(Calendar.MILLISECOND) * 10 ^ 6)
            .atZone(ZoneId.of("UTC"));
    DateTimeFormatter formatter = DateTimeFormatter.ISO_DATE_TIME;
    String result = formatter.format(d);
    return result.substring(0, result.length() - 12) + "Z";
  }

  public static Integer getColumnIndexFromTable(XSSFTable table, String columnName) {
    List<XSSFTableColumn> columns = table.getColumns();
    for (XSSFTableColumn c : columns) {
      if (columnName.equals(c.getName())) {
        return c.getColumnIndex() + table.getStartColIndex();
      }
    }
    return null;
  }

  public static void copyColumn(
      String fromSheetName, int fromColumn, XSSFSheet toSheet, int toColumn) {
    XSSFSheet fromSheet = toSheet.getWorkbook().getSheet(fromSheetName);
    Map<Integer, XSSFCellStyle> styleMap = new HashMap<>();
    JSONArray potentialErrors = new JSONArray();
    // Make space for the column that will be copied
    int lastFilledColumn = getLastFilledColumn(toSheet);
    log.info("Copy column " + toColumn + " with last filled " + lastFilledColumn);
    // Only shift columns if required
    if (lastFilledColumn > toColumn) toSheet.shiftColumns(toColumn + 1, lastFilledColumn, 1);
    for (int i = 0; i <= fromSheet.getLastRowNum(); i++) {
      XSSFRow r = fromSheet.getRow(i);
      if (r == null) continue;
      XSSFCell c = r.getCell(fromColumn);
      if (c == null) continue;
      XSSFCell toCell =
          Exporter.findAndCreateCellIfRequired(
              toSheet.getWorkbook(), toSheet.getSheetName(), i, toColumn + 1, potentialErrors);
      copyCell(c, toCell, styleMap);
    }
  }

  public static void removeColumn(Sheet sheet, int column) {
    int lastFilledColumn = getLastFilledColumn(sheet);
    log.info("Shifting " + column + " last filled + " + lastFilledColumn);
    for (Row row : sheet) {
      Cell cell = row.getCell(column);
      if (cell != null) {
        row.removeCell(cell);
      }
    }
    // Only column shift if the column removed is not the last column in the sheet¨
    // as it then throws a "firstMovedIndex, lastMovedIndex out of order" IllegalArgumentException
    if (column + 1 >= lastFilledColumn) return;
    log.fine("Removing column " + column + ", last filled " + getLastFilledColumn(sheet));
    if (column == 0) column = 1;
    sheet.shiftColumns(column, getLastFilledColumn(sheet), -1);
  }

  public static void removeCalcChain(XSSFWorkbook workbook) throws Exception {
    CalculationChain calcchain = workbook.getCalculationChain();
    Method removeRelation =
        POIXMLDocumentPart.class.getDeclaredMethod("removeRelation", POIXMLDocumentPart.class);
    removeRelation.setAccessible(true);
    removeRelation.invoke(workbook, calcchain);
  }

  // Used for column shifting
  private static int getLastFilledColumn(Sheet sheet) {
    int result = 0;
    for (Row row : sheet) {
      if (row.getLastCellNum() > result) result = row.getLastCellNum();
    }
    return result;
  }

  public static void copyCell(
      XSSFCell oldCell, XSSFCell newCell, Map<Integer, XSSFCellStyle> styleMap) {
    if (styleMap != null) {
      if (oldCell.getSheet().getWorkbook() == newCell.getSheet().getWorkbook()) {
        newCell.setCellStyle(oldCell.getCellStyle());
      } else {
        int stHashCode = oldCell.getCellStyle().hashCode();
        XSSFCellStyle newCellStyle = styleMap.get(stHashCode);
        if (newCellStyle == null) {
          newCellStyle = newCell.getSheet().getWorkbook().createCellStyle();
          newCellStyle.cloneStyleFrom(oldCell.getCellStyle());
          styleMap.put(stHashCode, newCellStyle);
        }
        newCell.setCellStyle(newCellStyle);
      }
    }
    switch (oldCell.getCellType()) {
      case STRING:
        newCell.setCellValue(oldCell.getStringCellValue());
        break;
      case NUMERIC:
        newCell.setCellValue(oldCell.getNumericCellValue());
        break;
      case BLANK:
        newCell.setCellType(CellType.BLANK);
        break;
      case BOOLEAN:
        newCell.setCellValue(oldCell.getBooleanCellValue());
        break;
      case ERROR:
        newCell.setCellErrorValue(oldCell.getErrorCellValue());
        break;
      case FORMULA:
        newCell.setCellFormula(oldCell.getCellFormula());
        break;
      default:
        break;
    }
  }

  public static List<XSSFTable> getAllTables(XSSFWorkbook workbook) {
    List<XSSFTable> allTables = new ArrayList<>();
    int numberOfSheets = workbook.getNumberOfSheets();
    for (int sheetIdx = 0; sheetIdx < numberOfSheets; sheetIdx++) {
      XSSFSheet sheet = workbook.getSheetAt(sheetIdx);
      List<XSSFTable> tables = sheet.getTables();
      allTables.addAll(tables);
    }
    return allTables;
  }

  public static String readFileAsString(String filePath, boolean absolutePath)
      throws UnsupportedEncodingException, IOException {
    String fullPath;
    if (!absolutePath) {
      final String currentDir = System.getProperty("user.dir");
      fullPath = currentDir + "/" + filePath;
    } else fullPath = filePath;
    return new String(Files.readAllBytes(Paths.get(fullPath)), StandardCharsets.UTF_8);
  }

  public static void writeOutWorkbookAsFile(Template template, File out) throws IOException {
    if (!out.exists()) {
      boolean exists = out.createNewFile();
      if (exists) log.warning("Overwrote " + out.getCanonicalPath());
    }
    try (FileOutputStream stream = new FileOutputStream(out)) {
      template.workbook.write(stream);
      stream.flush();
    } catch (IOException e) {
      throw e;
    } finally {
      template.workbook.close();
    }
  }

  public static List<String> diffTwoWorkbooksAndReturnErrors(Workbook expected, Workbook actual) {
    List<String> errors = new ArrayList<>();
    if (expected.getNumberOfSheets() != actual.getNumberOfSheets())
      errors.add(
          "Number of sheets differ, expected ["
              + expected.getNumberOfSheets()
              + "] does not match actual ["
              + actual.getNumberOfSheets()
              + "]");
    for (int i = 0; i < expected.getNumberOfSheets(); i++) {
      XSSFSheet es = (XSSFSheet) expected.getSheetAt(i), as = (XSSFSheet) actual.getSheetAt(i);
      if (!es.getSheetName().equals(as.getSheetName()))
        errors.add(
            "Sheet name expected ["
                + es.getSheetName()
                + "] does not match actual ["
                + as.getSheetName()
                + "]");
      if (es.getLastRowNum() > as.getLastRowNum())
        errors.add(
            "Number of rows in sheet ["
                + es.getSheetName()
                + "] differ, expected ["
                + es.getLastRowNum()
                + "] does not match actual ["
                + as.getLastRowNum()
                + "]");
      for (int r = 0; r <= es.getLastRowNum(); r++) {
        Row er = es.getRow(r), ar = as.getRow(r);
        if (er == null) continue;
        for (Cell ec : er) {
          if (ar == null) {
            errors.add(
                "Row ["
                    + (r + 1)
                    + "] exists in expected but not in actual sheet ["
                    + es.getSheetName()
                    + "]");
            continue;
          }
          Cell ac = ar.getCell(ec.getColumnIndex());
          String eValue = getCellValueAsString(ec), aValue = getCellValueAsString(ac);
          if (!eValue.equals(aValue))
            errors.add(
                "Value in sheet ["
                    + es.getSheetName()
                    + "] for cell ["
                    + ec.getAddress()
                    + "] differ, expected ["
                    + eValue
                    + "] does not match actual ["
                    + aValue
                    + "]");
        }
      }

      for (XSSFTable et : es.getTables()) {
        boolean match = false;
        for (XSSFTable at : as.getTables()) {
          if (et.getName() != null && et.getName().equals(at.getName())) {
            match = true;
            if (et.getRowCount() != at.getRowCount()) {
              errors.add(
                  "Sheet ["
                      + es.getSheetName()
                      + "] has a table ["
                      + et.getName()
                      + "] with row count ["
                      + et.getRowCount()
                      + "] but output has row count ["
                      + at.getRowCount()
                      + "]");
            }
          }
        }
        if (!match) {
          errors.add(
              "Sheet ["
                  + es.getSheetName()
                  + "] has a table ["
                  + et.getName()
                  + "] that was not found in the output");
        }
      }
    }
    return errors;
  }

  private static final DataFormatter dataFormatter = new DataFormatter(Locale.GERMAN);

  public static String getCellValueAsString(Cell cell) {
    if (cell == null) return "";
    else if (cell.getCellType() == CellType.FORMULA) {
      if (cell.getCachedFormulaResultType() == CellType.STRING) return cell.getStringCellValue();
      else if (cell.getCachedFormulaResultType() == CellType.NUMERIC) {
        String formatString = cell.getCellStyle().getDataFormatString();
        int formatIndex = cell.getCellStyle().getDataFormat();
        return dataFormatter.formatRawCellContents(
            cell.getNumericCellValue(), formatIndex, formatString);
      } else if (cell.getCachedFormulaResultType() == CellType.BLANK
          || cell.getCachedFormulaResultType() == CellType._NONE) return "";
      else if (cell.getCachedFormulaResultType() == CellType.BOOLEAN)
        return Boolean.toString(cell.getBooleanCellValue());
      else if (cell.getCachedFormulaResultType() == CellType.ERROR)
        return FormulaError.forInt(cell.getErrorCellValue()).getString();
      else
        throw new RuntimeException(
            "Cell ["
                + cell.getAddress().toString()
                + "] with type ["
                + cell.getCellType().toString()
                + "] and value ["
                + cell.getStringCellValue()
                + "]");
    } else {
      try {
        return dataFormatter.formatCellValue(cell);
      } catch (IllegalArgumentException e) {
        log.warning(
            "Formatting "
                + cell.getAddress().formatAsString()
                + " with format "
                + cell.getCellStyle().getDataFormatString()
                + " with value "
                + Utils.getSimpleCellValue(cell));
        return "";
      }
    }
  }

  private static String getSimpleCellValue(Cell cell) {
    if (cell == null) return null;
    switch (cell.getCellType()) {
      case BOOLEAN:
        if (cell.getBooleanCellValue()) return "true";
        else return "false";
      case NUMERIC:
        return Double.toString(cell.getNumericCellValue());
      case STRING:
        return cell.getStringCellValue();
      case BLANK:
        return "";
      case ERROR:
        return "Error in cell";
      default:
        return "Error: Unknown cell type";
    }
  }

  public static void copyFromInputToOutput(InputStream in, OutputStream out) {
    copyFromInputToOutput(in, out, 8192); // Default buffer size
  }

  /**
   * Copy from one stream to another, with a buffer
   *
   * @param in
   * @param out
   */
  public static void copyFromInputToOutput(
      InputStream in, OutputStream out, int bufferSizeInBytes) {
    try {
      byte[] buffer = new byte[bufferSizeInBytes]; // Measure in bytes, e.g., 8192 = 8kB
      int len;
      while ((len = in.read(buffer)) != -1) {
        out.write(buffer, 0, len);
      }
      out.flush();
    } catch (Exception e) {
      log.warning("Error copying stream: " + e.getMessage());
    } finally {
      try {
        out.close();
      } catch (IOException e) {
        log.warning("Error closing output stream: " + e.getMessage());
      }
    }
  }

  public static File saveContentsOfUrlAsTmpFile(String urlString) throws IOException {
    log.info("Will download data from [" + urlString + "]");
    File tmp = File.createTempFile("exported-pdf-" + new Random().nextInt(1000000), ".pdf");
    try (FileOutputStream out = new FileOutputStream(tmp)) {
      Utils.copyFromInputToOutput(getInputStreamFromURL(urlString), out);
    } catch (IOException e) {
      throw e;
    }
    return tmp;
  }

  // Used for testing only
  public static InputStream getInputStreamFromURL(String urlString) throws IOException {
    URL url = new URL(urlString);
    HttpURLConnection conn = (HttpURLConnection) url.openConnection();
    HttpURLConnection.setFollowRedirects(true);
    return conn.getInputStream();
  }

  public static CellColorRepresentation fromCell(Cell cell) {
    XSSFCellStyle st = ((XSSFCellStyle) cell.getCellStyle());
    XSSFColor color = st.getFillForegroundXSSFColor();
    if (color == null) return CellColorRepresentation.NONE; // No fill
    byte[] ar = st.getFillForegroundXSSFColor().getARGB(); // Alpha, Red,
    // Green, Blue
    if (ar == null || ar.length != 4)
      return CellColorRepresentation.NONE; // Not valid byte array for RGB
    byte r = ar[1], g = ar[2], b = ar[3]; // a = ar[0],
    // log.info(cell.getAddress() + " color values: r["+r+"]g["+g+"]b["+b+"]"); // alpha ["+a+"]");
    // // Use this to easily
    // log colors for each cell

    CellColorRepresentation result = CellColorRepresentation.NONE;
    if (r == (byte) 0 && g == (byte) -80 && b == (byte) 80)
      result = CellColorRepresentation.INPUT; // Excel 2016 Mac
    else if (r == (byte) 0 && g == (byte) -128 && b == (byte) 0)
      result = CellColorRepresentation.INPUT; // Excel 2011 Mac
    else if (r == (byte) 0 && g == (byte) -1 && b == (byte) 0)
      result = CellColorRepresentation.INPUT; // Google Sheets
    else if (r == (byte) -1 && g == (byte) 0 && b == (byte) 0)
      result = CellColorRepresentation.OUTPUT; // Excel 2016 Mac + Google
    // Sheets
    else if (r == (byte) 0 && g == (byte) 112 && b == (byte) -64)
      result = CellColorRepresentation.AGGREGATE; // Excel 2016 Mac
    else if (r == (byte) 51 && g == (byte) 102 && b == (byte) -1)
      result = CellColorRepresentation.AGGREGATE; // Excel 2011 Mac
    else if (r == (byte) 0 && g == (byte) 0 && b == (byte) -1)
      result = CellColorRepresentation.AGGREGATE; // Google Sheets
    else if (r == (byte) 112 && g == (byte) 48 && b == (byte) -96)
      result = CellColorRepresentation.METADATA; // Excel 2016 Mac
    else if (r == (byte) 102 && g == (byte) 0 && b == (byte) 102)
      result = CellColorRepresentation.METADATA; // Excel 2011 Mac
    else if (r == (byte) -103 && g == (byte) 0 && b == (byte) -1)
      result = CellColorRepresentation.METADATA; // Google Sheets
    else if (r == (byte) -1 && g == (byte) 0 && b == (byte) -1)
      result = CellColorRepresentation.METADATA; // Google Sheets
    else if (r == (byte) -1 && g == (byte) -1 && b == (byte) 0)
      result = CellColorRepresentation.ACTION; // Excel 2016 Mac
    // log.info(cell.getAddress()+" match ["+result+"]");
    return result;
  }

  enum CellColorRepresentation {
    NONE,
    INPUT,
    OUTPUT,
    AGGREGATE,
    ACTION,
    METADATA
  }

  enum Direction {
    LEFT,
    RIGHT,
    DOWN,
    UP
  }

  protected static Cell move(Cell startingCell, Direction... direction) {
    if (startingCell == null) return null;
    Sheet sheet = startingCell.getSheet();
    int row = startingCell.getRow().getRowNum();
    int col = startingCell.getColumnIndex();
    // log.info("Looking at row " + row + " col " + col);
    for (Direction d : direction) {
      if (d == Direction.LEFT) col--;
      else if (d == Direction.RIGHT) col++;
      else if (d == Direction.DOWN) row++;
      else if (d == Direction.UP) row--;
    }

    // log.info("Now at row " + row + " col " + col);
    Row newRow = sheet.getRow(row);
    if (newRow == null) return null; // No more rows
    else if (col >= 0) return newRow.getCell(col);
    else return null;
  }

  /**
   * POI doesn't support to modify regions that data validation applies to. Hence, we need to go
   * down in the underlying XML
   *
   * @param sheet The sheet to operate on
   * @param oldRegion The original region of the data validation
   * @param additionalRegion The region to expand the data validation with
   * @return
   */
  public static boolean expandDataValidationRegion(
      final XSSFSheet sheet,
      final CellRangeAddressList oldRegion,
      final CellRangeAddressList additionalRegion) {
    if (sheet == null) throw new UnsupportedOperationException("Sheet cannot be null");
    if (oldRegion == null) throw new UnsupportedOperationException("oldRegion cannot be null");
    if (additionalRegion == null)
      throw new UnsupportedOperationException("additionalRegion cannot be null");

    final List<String> oldSqref = convertSqref(oldRegion);
    try {
      Field fWorksheet = XSSFSheet.class.getDeclaredField("worksheet");
      fWorksheet.setAccessible(true);
      CTWorksheet worksheet = (CTWorksheet) fWorksheet.get(sheet);
      CTDataValidations dataValidations = worksheet.getDataValidations();
      if (dataValidations == null) return false;

      for (int i = 0; i < dataValidations.getCount(); i++) {
        CTDataValidation dv = dataValidations.getDataValidationArray(i);
        List<String> sqref = new ArrayList<>(dv.getSqref());
        // Match the right data validation based on the range it applies to
        if (equalsSqref(sqref, oldSqref)) {
          sqref.addAll(convertSqref(additionalRegion));
          dv.setSqref(sqref);
          StringJoiner joiner = new StringJoiner(",");
          sqref.forEach(item -> joiner.add(item));
          log.fine("New data validation region: " + joiner);
          dataValidations.setDataValidationArray(i, dv);
          return true;
        }
      }
    } catch (Exception e) {
      log.warning("Failed to expand data validation region: " + e.getMessage());
    }
    return false;
  }

  private static List<String> convertSqref(final CellRangeAddressList region) {
    List<String> sqref = new ArrayList<>();
    for (CellRangeAddress range : region.getCellRangeAddresses()) {
      sqref.add(range.formatAsString());
    }
    return sqref;
  }

  public static boolean equalsSqref(final List<String> sqref1, final List<String> sqref2) {
    if (sqref1.size() != sqref2.size()) {
      return false;
    }
    Collections.sort(sqref1);
    Collections.sort(sqref2);
    final int size = sqref1.size();
    for (int i = 0; i < size; i++) {
      if (!sqref1.get(i).equals(sqref2.get(i))) {
        return false;
      }
    }
    return true;
  }
}
