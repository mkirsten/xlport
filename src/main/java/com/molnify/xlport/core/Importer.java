package com.molnify.xlport.core;

import static com.molnify.xlport.core.Utils.*;

import java.lang.reflect.Array;
import java.util.*;
import java.util.Map.Entry;
import java.util.logging.Logger;
import org.apache.poi.ss.formula.eval.ErrorEval;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFName;
import org.apache.poi.xssf.usermodel.XSSFTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;

/**
 * Extracts structured JSON data from Excel workbooks (import).
 *
 * <p>Supports named ranges (properties), tables, wildcard extraction, multi-sheet imports,
 * and formula evaluation.
 *
 * <p>Usage:
 * <pre>{@code
 * XSSFWorkbook workbook = new XSSFWorkbook(new File("data.xlsx"));
 * JSONArray errors = new JSONArray();
 * JSONObject result = Importer.importAllData(workbook, errors, true);
 * }</pre>
 */
public class Importer {

  private static final Logger log = Logger.getLogger(Importer.class.getName());

  // Why public? So we can refer to these from tests
  public static final String PROPERTIES = "properties";
  public static final String TABLES = "tables";

  /**
   * Imports all named ranges (properties) and tables from the workbook.
   *
   * @param workbook the Excel workbook to extract data from
   * @param potentialErrors accumulator for any errors encountered during import
   * @param evaluate if true, evaluate formulas before reading cell values
   * @return JSON object with "properties" and "tables" keys
   */
  public static JSONObject importAllData(
      Workbook workbook, JSONArray potentialErrors, boolean evaluate) {
    return importData(
        new JSONObject()
            .put(PROPERTIES, new JSONArray().put("*"))
            .put(TABLES, new JSONArray().put("*")),
        (XSSFWorkbook) workbook,
        potentialErrors,
        true);
  }

  /**
   * Imports data from the workbook according to the request specification.
   *
   * @param request JSON specifying which "properties" (named ranges) and "tables" to extract.
   *     Use {@code ["*"]} as a wildcard to extract all.
   * @param workbook the Excel workbook to extract data from
   * @param potentialErrors accumulator for any errors encountered during import
   * @param evaluate if true, evaluate formulas before reading cell values
   * @return JSON object with "properties" and/or "tables" keys matching the request
   */
  public static JSONObject importData(
      JSONObject request, XSSFWorkbook workbook, JSONArray potentialErrors, boolean evaluate) {
    FormulaEvaluator evaluator = null;
    if (evaluate) {
      evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      evaluator.setIgnoreMissingWorkbooks(true);
    }

    JSONObject result = new JSONObject();

    // Two cases below: PROPERTIES and/or TABLES
    if (request.has(PROPERTIES)) {
      Set<String> props = new HashSet<>();
      Object p = request.get(PROPERTIES);
      if (p instanceof JSONArray) {
        for (Object o : ((JSONArray) p)) {
          props.add((String) o);
        }
      } else if (p instanceof JSONObject) {
        props.add(((JSONObject) p).toString());
      } else {
        potentialErrors.put(
            "Properties is specified incorrectly. It should either be a JSON array or a single JSON"
                + " object");
        return result;
      }

      // Wildcard support
      if (props.size() == 1
          && props.toArray()[0] instanceof String
          && "*".equals(props.toArray()[0])) {
        props = new HashSet<>();
        for (XSSFName n : workbook.getAllNames()) {
          props.add(n.getNameName());
        }
      }

      // Actual processing
      JSONObject resultProps = new JSONObject();
      result.put(PROPERTIES, resultProps);
      for (Object o : props) {
        if (o instanceof String) {
          String key = (String) o;
          for (XSSFName name : workbook.getNames(key)) {
            JSONObject jsonObjectToPlaceContent = resultProps;
            if (name == null) continue;
            // If the property is sheet scoped, put it in a JSONObject in the resultProps
            if (name.getSheetIndex() != -1) {
              if (!resultProps.has(name.getSheetName()))
                resultProps.put(name.getSheetName(), new JSONObject());
              jsonObjectToPlaceContent = resultProps.getJSONObject(name.getSheetName());
            }
            String ref = name.getRefersToFormula();
            // log.info("Will match the following ["+key+"] -> ["+ref+"]");
            try {
              Cell c = getNullSafeCellFromReference(new CellReference(ref), workbook);
              if (evaluator == null)
                putCellContentsInJSON(
                    c, c == null ? CellType._NONE : c.getCellType(), jsonObjectToPlaceContent, key);
              else putCellContentsInJSON(c, jsonObjectToPlaceContent, key, evaluator);
            } catch (java.lang.IllegalArgumentException e) {
              log.info(
                  "Ranges not supported yet. Try using a table instead for ["
                      + key
                      + "] -> ["
                      + ref
                      + "]");
            }
          }
        }
      }
    }

    if (request.has(TABLES)) {
      JSONArray tables = null;
      Object t = request.get(TABLES);
      if (t instanceof JSONArray) tables = (JSONArray) t;
      else if (t instanceof JSONObject) tables = new JSONArray().put(t);
      else {
        potentialErrors.put(
            "Tables is specified incorrectly. It should either be a JSON array or a single JSON"
                + " object");
        return result;
      }

      // Wildcard support
      if (tables.length() == 1
          && tables.get(0) instanceof String
          && "*".equals(tables.getString(0))) {
        tables.remove(0);
        for (XSSFTable n : getAllTables(workbook)) {
          tables.put(n.getName());
        }
      }

      // Actual processing
      JSONObject resultTables = new JSONObject();
      result.put(TABLES, resultTables);
      for (Object o : tables) {
        String tableName = null;
        List<String> headerNames = null;
        if (o instanceof JSONObject) { // Some headers to include
          JSONObject tableSpec = (JSONObject) o;
          tableName = (String) tableSpec.keySet().toArray()[0];
          headerNames = getStringsInJSONArray(tableSpec.getJSONArray(tableName));
        } else if (o instanceof String) { // Just table name, include all headers
          tableName = (String) o;
        }

        JSONObject jsonObjectToPlaceContent = resultTables;
        // If the tableName fetched is actually a collapsed name (i.e., can be expanded), then it
        // should be placed differently in the JSONObject
        String[] expandedTableName = Utils.expandName(tableName);
        if (expandedTableName.length == 2) {
          String sheetName = expandedTableName[0];
          String displayTableName = expandedTableName[1];
          if (!resultTables.has(sheetName)) resultTables.put(sheetName, new JSONObject());
          jsonObjectToPlaceContent = resultTables.getJSONObject(sheetName);
          JSONArray array =
              getExcelTableAsArrayOfJSONObjects(tableName, headerNames, workbook, evaluator);
          jsonObjectToPlaceContent.put(displayTableName, array);
        } else {
          JSONArray array =
              getExcelTableAsArrayOfJSONObjects(tableName, headerNames, workbook, evaluator);
          jsonObjectToPlaceContent.put(tableName, array);
        }
      }
    }
    return result;
  }

  /**
   * If headerNames are empty is is null, this method will automatically pick up all headers/columns
   */
  private static JSONArray getExcelTableAsArrayOfJSONObjects(
      String tableName,
      List<String> headerNames,
      XSSFWorkbook workbook,
      FormulaEvaluator evaluator) {
    JSONArray result = new JSONArray();
    List<XSSFTable> tables = getAllTables(workbook);
    for (XSSFTable t : tables) {
      if (tableName.equals(t.getDisplayName())) {
        // Build up a map between the name of a header in the table and its column
        Map<String, Integer> headerNameToColumn = new HashMap<String, Integer>();
        int startingRow = t.getStartCellReference().getRow();
        for (int col = t.getStartColIndex(); col <= t.getEndColIndex(); col++) {
          CellReference ref = new CellReference(t.getSheetName(), startingRow, col, true, true);
          Cell c = getNullSafeCellFromReference(ref, workbook);
          String headerName = getCellValueAsString(c);
          if (headerNames == null || headerNames.size() == 0 || headerNames.contains(headerName))
            headerNameToColumn.put(headerName, col);
        }

        // Now loop through the table and pick out relevant columns
        for (int row = startingRow + 1; row <= t.getEndRowIndex(); row++) {
          boolean rowEmtpySoFar = true;
          JSONObject dataForRow = new JSONObject();
          for (Entry<String, Integer> e : headerNameToColumn.entrySet()) {
            CellReference ref = new CellReference(t.getSheetName(), row, e.getValue(), true, true);
            Cell c = getNullSafeCellFromReference(ref, workbook);
            if (evaluator == null) {
              boolean empty =
                  putCellContentsInJSON(
                      c, c == null ? CellType._NONE : c.getCellType(), dataForRow, e.getKey());
              if (!empty) rowEmtpySoFar = false;
            } else {
              boolean empty = putCellContentsInJSON(c, dataForRow, e.getKey(), evaluator);
              if (!empty) rowEmtpySoFar = false;

              // Extract certain inputs as required, only when an evaluator is included
              if (c != null && fromCell(c) == CellColorRepresentation.INPUT) {
                Cell d = move(c, Direction.RIGHT);
                if (d != null) {
                  String contents = getCellValueAsString(d);
                  if (contents != null && contents.toLowerCase().contains("fileupload")) {
                    if (!dataForRow.isNull(e.getKey())) {
                      Object[] values =
                          (Object[])
                              getJSONCompatibleArrayFromString(dataForRow.get(e.getKey()), ",");

                      // Add both downsized and full-sized URLs
                      for (int i = 0; i < Array.getLength(values); i++) {
                        if (Array.get(values, i) == null) continue;
                        String url = Array.get(values, i).toString(), fullUrl = "";
                        int lastDot = url.lastIndexOf('.');
                        if (lastDot >= 0 && lastDot > url.lastIndexOf('/'))
                          fullUrl = url.substring(0, lastDot) + "__full_" + url.substring(lastDot);
                        else fullUrl = url + "__full_";

                        JSONObject o = new JSONObject();
                        o.put("url", url);
                        o.put("url_full", fullUrl);

                        Array.set(values, i, o);
                      }
                      dataForRow.put(e.getKey() + "_array", values);
                    } else dataForRow.put(e.getKey() + "_array", JSONObject.NULL);

                  } else if (contents != null && contents.toLowerCase().contains("multiple")) {
                    dataForRow.put(
                        e.getKey() + "_array",
                        getJSONCompatibleArrayFromString(dataForRow.get(e.getKey()), ";"));
                  }
                }
              }
            }
          }
          result.put(dataForRow);
          if (rowEmtpySoFar) {
            log.info(
                "Hey, I'm done exporting at row "
                    + row
                    + "/"
                    + t.getEndRowIndex()
                    + " for "
                    + tableName
                    + ". Checking out");
            break;
          }
        }
      }
    }
    return result;
  }

  private static Object getJSONCompatibleArrayFromString(Object o, String s) {
    if (!(o != null && o instanceof String)) return JSONObject.NULL;
    return Arrays.stream(((String) o).split(s)).map(String::trim).toArray();
  }

  private static List<String> getStringsInJSONArray(JSONArray array) {
    List<String> result = new ArrayList<String>();
    for (Object o : array) {
      if (o instanceof String) result.add((String) o);
      else log.info("JSONArray [" + array + "] did NOT contain string [" + o + "]");
    }
    return result;
  }

  /** Returns the cell on cellRef from workbook if it can be found, otherwise null */
  private static Cell getNullSafeCellFromReference(CellReference cellRef, Workbook workbook) {
    String sheetName = cellRef.getSheetName();
    try {
      Sheet s = workbook.getSheet(sheetName);
      if (s == null) return null;
      Row r = s.getRow(cellRef.getRow());
      if (r == null) return null;
      return r.getCell(cellRef.getCol());
    } catch (NullPointerException npe) {
      return null;
    }
  }

  /*
   * Why include type here and not only the cell? So we can call the method from
   * itself if the type is formula, and get the cached result
   */
  private static boolean putCellContentsInJSON(Cell c, CellType type, JSONObject obj, String key) {
    switch (type) {
      case _NONE:
      case BLANK:
        obj.put(key, JSONObject.NULL);
        return true;
      case STRING:
        obj.put(key, c.getStringCellValue());
        break;
      case BOOLEAN:
        obj.put(key, c.getBooleanCellValue());
        break;
      case ERROR:
        obj.put(key, ErrorEval.valueOf(c.getErrorCellValue()).getErrorString());
        break;
      case NUMERIC:
        if (DateUtil.isCellDateFormatted(c)) {
          Date date = DateUtil.getJavaDate(c.getNumericCellValue());
          String s = getAsISO8601String(date);
          obj.put(key, s);
        } else obj.put(key, c.getNumericCellValue());
        break;
      case FORMULA:
        putCellContentsInJSON(c, c.getCachedFormulaResultType(), obj, key);
        break;
    }
    return false;
  }

  /**
   * Returns true it the data was considered empty
   *
   * <p>This method also evaluates the cell
   */
  private static boolean putCellContentsInJSON(
      Cell c, JSONObject obj, String key, FormulaEvaluator evaluator) {
    CellType type = CellType._NONE;
    CellValue cellValue = null;
    try {
      if (c != null) {
        // log.info("Evaluating " + c.getAddress().formatAsString() + " for key " + key);
        try {
          cellValue = evaluator.evaluate(c);
        } catch (Throwable t) {
          log.warning("Failed to evaluate cell: " + t.getMessage());
        }
        if (cellValue != null) {
          type = cellValue.getCellType();
          // log.info("Evaluated: " + cellValue.formatAsString());
        } else {
          // log.info("Skipping");
        }
      }
      switch (type) {
        case _NONE:
        case BLANK:
          obj.put(key, JSONObject.NULL);
          return true;
        case STRING:
          String str = cellValue.getStringValue();
          obj.put(key, str);
          return str == null || str.length() == 0;
        case BOOLEAN:
          obj.put(key, cellValue.getBooleanValue());
          return false;
        case ERROR:
          obj.put(key, ErrorEval.valueOf(cellValue.getErrorValue()).getErrorString());
          return true;
        case NUMERIC:
          try {
            if (DateUtil.isCellDateFormatted(c)) { // Will fail if cell is not NUMERIC
              Date date = DateUtil.getJavaDate(cellValue.getNumberValue());
              String s = getAsISO8601String(date);
              obj.put(key, s);
            } else obj.put(key, cellValue.getNumberValue());
          } catch (IllegalStateException e) {
            obj.put(key, cellValue.getNumberValue());
          }
          return false;
        case FORMULA:
          log.warning(
              "This should NEVER happen " + key + " cell " + c.getAddress().formatAsString());
        default:
          return true;
      }
    } catch (Throwable t) {
      obj.put(key, "#ERROR");
      log.warning("Error reading value from cell: " + c.getAddress().formatAsString() + " - " + t.getMessage());
      return true;
    }
  }
}
