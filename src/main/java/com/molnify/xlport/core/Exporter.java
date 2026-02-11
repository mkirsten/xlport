package com.molnify.xlport.core;

import java.util.*;
import java.util.List;
import java.util.logging.Logger;
import org.apache.poi.ss.SpreadsheetVersion;
import org.apache.poi.ss.formula.eval.NotImplementedFunctionException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.*;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;

/**
 * Populates an Excel template with JSON data (export).
 *
 * <p>Supports named ranges (single-cell properties), tables, multi-sheet templates, formulas,
 * data validation, conditional formatting, and cell-level formatting.
 *
 * <p>Usage:
 * <pre>{@code
 * Template template = TemplateManager.getTemplate("my-template.xlsx");
 * JSONObject data = new JSONObject(jsonString);
 * JSONArray errors = new JSONArray();
 * Exporter.exportToExcel(data, template, errors, true);
 * }</pre>
 */
public class Exporter {
  private static final Logger log = Logger.getLogger(Exporter.class.getName());

  private static final String SHEET_TEMPLATE_NAME = "sheets";

  /**
   * Populates the given template workbook with data from the JSON object.
   *
   * @param data JSON object containing key-value pairs that map to named ranges and tables
   *     in the template
   * @param template the Excel template to populate (modified in place)
   * @param potentialErrors accumulator for any errors encountered during export
   * @param overwriteFormatting if true, replicate cell styles from the first data row in tables
   */
  public static void exportToExcel(
      JSONObject data, Template template, JSONArray potentialErrors, boolean overwriteFormatting) {
    try { // Surround fully with try/catch to avoid HTTP 500 in the servlet
      // These are the sheets that are template for new sheets, and hence will be removed
      Set<String> templateSheetsToRemove = new HashSet<>();
      // This is a map from sheet names to a map of (old) tableNames with an AreaReference for their
      // area
      Map<String, Set<XSSFTable>> sheetToTables = new HashMap<>();
      if (data.has(SHEET_TEMPLATE_NAME))
        processSheetTemplates(
            data, template, potentialErrors, templateSheetsToRemove, sheetToTables);
      // We're done with the template sheets now, so let's chuck them away
      for (String sheetToRemove : templateSheetsToRemove)
        template.workbook.removeSheetAt(template.workbook.getSheetIndex(sheetToRemove));
      // All sheet template have been processed here, so remove it so the resulting JSON is easier
      // to read
      data.remove(SHEET_TEMPLATE_NAME);
      // System.out.println(data);
      // System.out.println(template.toString());

      // We do this to ensure keys that point to tables that have formulas are processed last,
      // in order to get all the data that they may look at be populated first
      LinkedList<String> theKeys = new LinkedList<>();
      for (String key : data.keySet()) {
        if (template.items.get(key) != null && template.items.get(key).isTableAndHasFormulas())
          theKeys.addLast(key);
        else theKeys.addFirst(key);
      }

      // Here is the main loop through all the keys in the request. Now ordered based on the
      // prioritization above
      for (String key : theKeys) {
        // This is the case where there is data for a specific sheet (with sheet scoped properties)
        if (data.get(key) instanceof JSONObject
            && (data.getJSONObject(key).has("_xlport_metadata")
                && "sheet".equals(data.getJSONObject(key).getString("_xlport_metadata")))) {
          for (String subKey : ((JSONObject) data.get(key)).keySet()) {
            processThisKey(
                data.getJSONObject(key),
                template,
                potentialErrors,
                overwriteFormatting,
                subKey,
                key);
          }
        }
        // ...and this is the normal case, with globally specified properties
        else processThisKey(data, template, potentialErrors, overwriteFormatting, key, null);
      }
      FormulaEvaluator evaluator = template.workbook.getCreationHelper().createFormulaEvaluator();
      template.workbook.setForceFormulaRecalculation(
          true); // Will ask Excel to evaluate when opened
      evaluator.evaluateAll();
    } catch (Throwable t) {
      log.warning("Export failed: " + t.getMessage());
      potentialErrors.put(t.getMessage());
      if (t.getCause() != null) {
        String additionalMessage = t.getCause().getMessage();
        if (t.getCause() instanceof NotImplementedFunctionException) {
          if (additionalMessage != null)
            additionalMessage = additionalMessage.replace("_xlfn.", "");
          additionalMessage = "Not implemented yet in xlPort: " + additionalMessage;
        }
        potentialErrors.put(additionalMessage);
      }
    }
  }

  private static void processSheetTemplates(
      JSONObject data,
      Template template,
      JSONArray potentialErrors,
      Set<String> templateSheetsToRemove,
      Map<String, Set<XSSFTable>> sheetToTables) {
    // First of all, let's find out which Sheets that are templates
    JSONArray templates = data.getJSONArray(SHEET_TEMPLATE_NAME);
    for (Object t : templates) {
      if (t instanceof JSONObject) {
        JSONObject tj = (JSONObject) t;
        if (tj.has("fromTemplateSheet") && tj.get("fromTemplateSheet") instanceof String)
          templateSheetsToRemove.add(tj.getString("fromTemplateSheet"));
      }
    }

    // Tables are buggy to copy, so let's take them away, save their metadata, and recreate them in
    // the new sheets
    for (String sheetName : templateSheetsToRemove) {
      log.info("Checking a sheet for tables... [" + sheetName + "]");
      if (sheetToTables.get(sheetName) == null)
        sheetToTables.put(
            sheetName,
            new HashSet<>()); // This avoid null check later, since all template sheets have a map
      for (XSSFTable t : template.workbook.getSheet(sheetName).getTables()) {
        t.getXSSFSheet().removeTable(t); // Weird call, but this is how it looks :-)
        sheetToTables.get(sheetName).add(t);
        log.info(
            "Copying data about a table that will later be removed ["
                + t.getName()
                + "] that is in the template sheet ["
                + sheetName
                + "]");
      }
    }

    // If there will be sheet additions (and removals), we need to handle named ranges on the old
    // and new sheets
    Map<String, Map<String, String>> savedNamedRangesByOriginalSheetName = new HashMap<>();
    if (templateSheetsToRemove.size() > 0) {
      List<Name> namesToRemove = new ArrayList<>();
      for (Name n : template.workbook.getAllNames()) {
        if (n.getRefersToFormula() == null || n.isFunctionName()) continue;
        String formula = n.getRefersToFormula();
        // Could throw IllegalArgument maybe if named range does not contain sheet name in formula
        // and is global
        String sheetName = null;
        try {
          sheetName =
              formula.indexOf("!") > 0
                  ? formula.substring(0, formula.indexOf("!"))
                  : n.getSheetName();
        } catch (IllegalArgumentException e) {
          log.info("Broken ref for " + n.getNameName() + ", will remove it");
          namesToRemove.add(n);
        }
        // This name will be deleted as part of copy, we need to create a new and scope it
        if (templateSheetsToRemove.contains(sheetName)) {
          if (!savedNamedRangesByOriginalSheetName.containsKey(n.getSheetName()))
            savedNamedRangesByOriginalSheetName.put(n.getSheetName(), new HashMap<>());
          Map<String, String> nameToFormula =
              savedNamedRangesByOriginalSheetName.get(n.getSheetName());
          if (formula.contains("!")) formula = formula.substring(formula.indexOf("!") + 1);
          nameToFormula.put(n.getNameName(), formula);
          namesToRemove.add(n);
          // System.out.println(n.getNameName() + " : " + formula);
        }
      }

      // Now we have saved away the names and are good to go to use the templates to create the new
      // sheets
      for (Object o : templates) {
        JSONObject sheetTemplate = (JSONObject) o;
        String toSheetName = sheetTemplate.getString("name");
        String fromTemplateSheetName = sheetTemplate.getString("fromTemplateSheet");
        if (template.workbook.getSheet(fromTemplateSheetName) == null) {
          potentialErrors.put(
              new JSONObject()
                  .put(
                      "message",
                      "template does not contain sheet ["
                          + fromTemplateSheetName
                          + "] so can't create ["
                          + toSheetName
                          + "]"));
          continue;
        }
        // JSONObject newData = sheetTemplate.getJSONObject("data");
        int indexToCopy = template.workbook.getSheetIndex(fromTemplateSheetName);

        // Copy the template sheet and name it with its new name. Make sure to position it correctly
        // in the workbook
        XSSFSheet copySheet = template.workbook.cloneSheet(indexToCopy, toSheetName);
        template.workbook.setSheetOrder(toSheetName, indexToCopy);

        // Add back named ranges from the template here and scope them to the sheet
        Map<String, String> m = savedNamedRangesByOriginalSheetName.get(fromTemplateSheetName);
        if (m != null)
          for (Map.Entry<String, String> nameToMigrate : m.entrySet()) {
            // log.info("Adding back " + nameToMigrate.getKey() + " to sheet " + toSheetName + "
            // value " + nameToMigrate.getValue());
            // Escape complex sheet names
            String referenceSheetName;
            if ((toSheetName.contains("-")
                    || toSheetName.contains("+")
                    || toSheetName.contains(" "))
                && !toSheetName.contains("'")) {
              referenceSheetName = "'" + toSheetName + "'";
            } else referenceSheetName = toSheetName;
            String reference = referenceSheetName + "!" + nameToMigrate.getValue();
            XSSFName localName = template.workbook.createName();
            localName.setSheetIndex(template.workbook.getSheetIndex(copySheet));
            localName.setNameName(nameToMigrate.getKey());
            localName.setRefersToFormula(reference);
            // When added to the template here, we can later just push into data to this reference
            template.addTemplateItemScoped(
                new TemplateItem(nameToMigrate.getKey(), reference, toSheetName), toSheetName);
          }

        // Fix tab color
        if (sheetTemplate.has("tabColor") && sheetTemplate.get("tabColor") instanceof String) {
          String tabColor = sheetTemplate.getString("tabColor");
          if (tabColor != null && tabColor.length() >= 6) {
            if (!tabColor.startsWith("#")) tabColor = "#" + tabColor;
            if (tabColor.length() == 7) {
              IndexedColorMap colorMap = template.workbook.getStylesSource().getIndexedColors();
              XSSFColor theColor = new XSSFColor(java.awt.Color.decode(tabColor), colorMap);
              copySheet.setTabColor(theColor);
            }
          }
        }

        // Add back the tables
        Set<XSSFTable> tablesInThisSheet = sheetToTables.get(fromTemplateSheetName);
        for (XSSFTable t : tablesInThisSheet) {
          log.info(
              "Adding back the table ["
                  + t.getName()
                  + "] that was in sheet ["
                  + fromTemplateSheetName
                  + "] and now will be in ["
                  + copySheet.getSheetName()
                  + "]");
          XSSFTable newT =
              createNewTableFromJSONSpecAndExistingTable(sheetTemplate, copySheet, t, true);
          template.addTemplateItem(new TemplateItem(newT));
        }

        // This will now copy over the data for the sheets into the root
        // Then the normal processing will continue, and use props (that may be locally scoped) as
        // well as tables and push in the data
        Utils.translateSheetTemplateSpecToRoot(sheetTemplate, data);
      } // End of template sheet
      for (Name n : namesToRemove) {
        template.workbook.removeName(n);
      }
    }
  }

  /**
   * @param sheetTemplate The JSON specification with data and columns
   * @param copySheet The sheet in which to create the table
   * @param t A XSSFTable, that has e.g., cell references, table name etc. (to use as a template)
   * @param collapseName Whether to collapse name from the sheet (i.e., __<sheetName>__<tableName>
   *     or just <tableName>
   * @return A XSSFTable created in the sheet copySheet
   */
  private static XSSFTable createNewTableFromJSONSpecAndExistingTable(
      JSONObject sheetTemplate, XSSFSheet copySheet, XSSFTable t, boolean collapseName) {
    // If this sheetTemplate has a spec for the table data for the columns
    int newColumnCount = t.getColumnCount();
    JSONObject sheetData = sheetTemplate.getJSONObject("data");
    if (sheetData.has(t.getName())
        && sheetData.get(t.getName()) instanceof JSONObject
        && ((JSONObject) sheetData.get(t.getName())).has("columns")
        && ((JSONObject) sheetData.get(t.getName())).get("columns") instanceof JSONArray) {
      // Idea here is to first remove all columns, and then add them back
      // It seems difficult to mangle around columns while removing some
      for (int i = 0; i <= (t.getEndColIndex() - t.getStartColIndex()); i++) {
        Utils.removeColumn(copySheet, t.getStartColIndex());
        // newT.removeColumn(i);
      }
      JSONArray columns = sheetData.getJSONObject(t.getName()).getJSONArray("columns");
      newColumnCount = columns.length();
      // This goes through the columns spec (from JSON) and adds all columns required
      // Two ways to spec 1) "name" and "fromTemplateColumn", and 2) just "name"
      // (2) is same as setting "name" and "fromTemplateColumn" to the same
      for (int i = columns.length() - 1; i >= 0; i--) {
        Object columnObj = columns.get(i);
        JSONObject c = (JSONObject) columnObj;
        String columnName = c.getString("name");
        String fromColumn;
        if (c.has("fromTemplateColumn")) fromColumn = c.getString("fromTemplateColumn");
        else fromColumn = columnName; // Support for convenience that only spec name
        log.info("Copying from column " + fromColumn + " into column " + columnName);
        int columnIndexToCopyFrom = Utils.getColumnIndexFromTable(t, fromColumn);
        Utils.copyColumn(
            t.getSheetName(), columnIndexToCopyFrom, copySheet, t.getStartColIndex() - 1);
        Cell cel =
            findAndCreateCellIfRequired(
                copySheet.getWorkbook(),
                copySheet.getSheetName(),
                t.getStartRowIndex(),
                t.getStartColIndex(),
                new JSONArray());
        cel.setCellValue(columnName);

        // If there is a format also specified, apply it to the first data row in the table
        // This is the cell right below the header. The style will then be copied to all data rows
        // (by later code)
        if (c.has("format") && c.get("format") instanceof String) {
          String format = c.getString("format");
          XSSFCellStyle cellStyle = copySheet.getWorkbook().createCellStyle();
          XSSFDataFormat dataFormat = copySheet.getWorkbook().createDataFormat();
          cellStyle.setDataFormat(dataFormat.getFormat(format));
          Cell oneDown =
              findAndCreateCellIfRequired(
                  copySheet.getWorkbook(),
                  copySheet.getSheetName(),
                  t.getStartRowIndex() + 1,
                  t.getStartColIndex(),
                  new JSONArray());
          oneDown.setCellStyle(cellStyle);
          log.info(
              "Applying format ["
                  + format
                  + "] to cell ["
                  + oneDown.getAddress().formatAsString()
                  + "]");
        }
      }
    }

    // Here the new teble is created
    int columnsInTable = newColumnCount, rowsInTable = t.getRowCount();
    AreaReference originalRef = t.getArea();
    // First cell is the same as in previous table
    String first = originalRef.getFirstCell().formatAsString(false);
    CellReference last =
        new CellReference(
            originalRef.getLastCell().getRow() + (rowsInTable - t.getRowCount()),
            originalRef.getLastCell().getCol() + (columnsInTable - t.getColumnCount()));
    AreaReference newTableReference =
        new AreaReference(new CellReference(first), last, SpreadsheetVersion.EXCEL2007);
    // Collapse name when dynamically creating table collapsed with sheetName
    String newTableName =
        collapseName
            ? Utils.collapseName(copySheet.getSheetName(), t.getDisplayName())
            : t.getDisplayName();
    log.info("Adding to template: " + newTableName + " with reference " + newTableReference);
    return createTableFromReference(copySheet, newTableReference, newTableName);
  }

  /**
   * This goes through the JSON data, picks up the data at key (and potentially sheet), and inserts
   * it in the Excel workbook (as given in the template)
   *
   * @param data The JSON data
   * @param template The template object, also including a reference to the Excel workbook. This is
   *     where data is inserted
   * @param potentialErrors This is a list of errors that the method can add to, to enable better
   *     debugging
   * @param overwriteFormatting If true, don't copy style that can be inferred (only applicable for
   *     the table case)
   * @param key The key to pick up from the JSON data
   * @param sheet Optionally set as a string. If set, the method will use the data from the
   *     JSONObject in "data" specified by "key", and use the sheet scoped template item in
   *     "template"
   */
  private static void processThisKey(
      JSONObject data,
      Template template,
      JSONArray potentialErrors,
      boolean overwriteFormatting,
      String key,
      String sheet) {
    // System.out.println("Key ["+key+"] sheet ["+sheet+"] data : " + data);
    /*log.info("template: " + template);
    log.info("template.items: " + template.items);
    log.info("template.sheetScopedItems: " + template.sheetScopedItems);
    log.info("sheet: " + sheet);
    */
    TemplateItem item;
    if (sheet == null) item = template.items.get(key);
    else {
      Map<String, TemplateItem> stringTemplateItemMap = template.sheetScopedItems.get(sheet);
      if (stringTemplateItemMap != null) item = stringTemplateItemMap.get(key);
      else {
        log.info("Could not find item " + key + " in templates");
        item = null;
      }
      // System.out.println("Found " + item + " for sheet ["+sheet + "] and key ["+key+"]");
    }
    // log.info("Processing: " + key + " in sheet " + sheet);
    if (item == null) {
      return;
    }
    String[] parts = item.reference.split("!");
    if (parts.length != 2)
      potentialErrors.put(
          "Template name ["
              + key
              + "] does not appear to match a cell (range), as the reference to it is ["
              + item.reference
              + "]");
    else {
      String sheetName = parts[0];
      // In the TemplateItem reference, the sheet names have been escaped. However, when using the
      // sheet
      // for other than reference, they should not be escaped (e.g., workbook.getSheet(sheetName)
      if (sheetName.startsWith("'")) sheetName = sheetName.substring(1, sheetName.length() - 1);
      // if(sheetName.indexOf(" ") > 0) sheetName = "\"" + sheetName + "\"";
      String cellName = parts[1];
      if (!cellName.contains(":")) { // Single cell case
        // log.info("Updating sheet ["+sheetName+"] cell ["+cellName+"] with value
        // ["+data.get(key)+"]");
        CellReference ref = new CellReference(cellName);
        XSSFCell c =
            findAndCreateCellIfRequired(
                template.workbook, sheetName, ref.getRow(), ref.getCol(), potentialErrors);
        if (c == null) return;

        try {
          // System.out.println("For sheet " + sheet + " data; " + data.get(key));
          insertDataFromJSONIntoCell(data.get(key), c);
        } catch (Exception e) {
          log.info(
              "Try to insert ["
                  + data.get(key)
                  + "] into cell ["
                  + c.getAddress().toString()
                  + "]");
          potentialErrors.put(e.getMessage());
        }

      } else { // Table case
        log.info(
            "Making magic for ["
                + key
                + "] with ref ["
                + cellName
                + "] in sheet ["
                + sheetName
                + "]");
        if (!item.isTable()) throw new IllegalStateException("This should be a table but is not");
        clearOutTableFromReference(template.workbook, item.reference);

        JSONArray array;
        // Option 1: The table is specified with both data and columns. In that case, fix the
        // columns first
        if (data.get(key) instanceof JSONObject
            && data.getJSONObject(key).has("data")
            && data.getJSONObject(key).has("columns")) {
          XSSFTable table = template.workbook.getTable(key);
          template.workbook.getSheet(sheetName).removeTable(table);
          XSSFTable newT =
              createNewTableFromJSONSpecAndExistingTable(
                  new JSONObject().put("data", data),
                  template.workbook.getSheet(sheetName),
                  table,
                  false);
          // We are actually already processing a template item, but this now needs to be replaced,
          // as the columns may have changed
          TemplateItem newItem = new TemplateItem(newT);
          template.addTemplateItem(newItem);
          item = newItem;
          array = data.getJSONObject(key).getJSONArray("data");
        }
        // Option 2: The table is just a JSONArray
        else array = data.getJSONArray(key);

        if (array.length() == 0) return;
        template.workbook.getTable(item.name).setDataRowCount(array.length());

        CellStyle[] styles = new CellStyle[item.getHeaders().size() + 1];
        ConditionalFormatting[] condFormatting =
            new ConditionalFormatting[item.getHeaders().size() + 1];
        log.info("Set up " + styles.length + " placeholders for formatting");

        for (int i = 0; i < array.length(); i++) {
          // log.info("Table round ["+i+"]");
          JSONObject o;
          try {
            o = array.getJSONObject(i);
          } catch (JSONException e) {
            potentialErrors.put(
                new JSONObject()
                    .put("status", "error")
                    .put(
                        "message",
                        "JSONArray ["
                            + key
                            + "] does not contain a JSONObject on position ["
                            + i
                            + "]. Skipping."));
            continue;
          }
          for (String iKey : item.getHeaders()) {
            // log.info("	Table key ["+iKey+"]");
            int[] rowAndCol = item.getRowAndColumnForHeader(iKey);
            if (rowAndCol == null) {
              log.info("		Could not find [" + key + "." + iKey + "] in template");
              continue;
            }
            int row = rowAndCol[0], col = rowAndCol[1];
            XSSFCell c =
                findAndCreateCellIfRequired(
                    template.workbook, sheetName, row + i + 1, col, potentialErrors);

            // Insert the data from the json into the workbook
            Object value;
            if (o.has(iKey)) {
              value = o.get(iKey);
              if (o.isNull(iKey)) value = null; // Unexpected treatment of null in the JSON
            } else value = item.getFormulaForHeader(iKey);

            boolean formattedAtCellLevel = false;
            try {
              formattedAtCellLevel = insertDataFromJSONIntoCell(value, c);
            } catch (Exception e) {
              log.info(
                  "Try to insert ["
                      + value
                      + "] into cell ["
                      + (c == null ? "unknown" : c.getAddress().toString())
                      + "]");
              potentialErrors.put(e.getMessage());
            }

            // If the overwriteFormatting flag is set, skip format write down.
            if (!overwriteFormatting) continue;

            // Below is logic to replicate CellStyle, DataValidation and Conditional Formatting
            int relCol = col - item.startingColumn;
            // For first row, copy a) CellStyle, b) DataValidation, and c) ConditionalFormatting
            if (i == 0) {
              // log.info("First row for ["+iKey+"] column ["+col+"], address
              // ["+c.getAddress().toString()+"]");
              if (c == null) continue;
              if (relCol >= styles.length) {
                log.warning(
                    "This is a bug with relCol ["
                        + relCol
                        + "] and styles.length ["
                        + styles.length
                        + "]");
                continue;
              }
              styles[relCol] = c.getCellStyle();
              for (DataValidation dv : c.getSheet().getDataValidations()) {
                for (CellRangeAddress cra : dv.getRegions().getCellRangeAddresses()) {
                  if (cra.isInRange(c)) {
                    Utils.expandDataValidationRegion(
                        c.getSheet(),
                        dv.getRegions(),
                        new CellRangeAddressList(row + 2, row + array.length(), col, col));
                  }
                }
              }
              SheetConditionalFormatting scf = c.getSheet().getSheetConditionalFormatting();
              for (int j = 0; j < scf.getNumConditionalFormattings(); j++) {
                ConditionalFormatting cf = scf.getConditionalFormattingAt(j);
                for (CellRangeAddress cra : cf.getFormattingRanges()) {
                  if (cra.isInRange(c)) condFormatting[relCol] = cf;
                }
              }
            } else { // If not first row, paste in the (a), (b) and (c) from above
              int nr = c.getAddress().getRow(), nc = c.getAddress().getColumn();
              if (formattedAtCellLevel) continue;

              if (styles[relCol] != null) c.setCellStyle(styles[relCol]);
              if (condFormatting[relCol] != null) {
                CellRangeAddress[] a = condFormatting[relCol].getFormattingRanges();
                CellRangeAddress[] aNew = new CellRangeAddress[a.length + 1];
                for (int j = 0; j < a.length; j++) {
                  aNew[j] = a[j];
                }
                aNew[a.length] = new CellRangeAddress(nr, nr, nc, nc);
                condFormatting[relCol].setFormattingRanges(aNew);
              }
            }
          }
        }
      }
      // End table case
    }
  }

  /** Clears out a table, excluding header + formula cells */
  public static void clearOutTableFromReference(Workbook workbook, String reference) {
    String[] sp = reference.split("!");
    if (sp.length != 2) {
      throw new IllegalStateException(
          "Cell reference should include sheet name [" + reference + "]");
    }
    String sheetName = sp[0];

    String[] cells = sp[1].split(":");
    if (cells.length != 2) {
      throw new IllegalStateException("Cell reference should be two cells [" + sp[1] + "]");
    }
    CellReference s = new CellReference(cells[0]), e = new CellReference(cells[1]);
    Sheet sheet = workbook.getSheet(sheetName);
    for (int row = s.getRow() + 1; row <= e.getRow(); row++)
      for (int col = s.getCol(); col <= e.getCol(); col++) {
        Row r = sheet.getRow(row);
        if (r == null) continue;
        Cell c = r.getCell(col);
        if (c == null) continue;
        if (c.getCellType() == CellType.FORMULA || c.getCellType() == CellType.BLANK) continue;
        else c.setCellValue("");
      }
  }

  public static XSSFTable createTableFromReference(
      XSSFSheet sheet, AreaReference ref, String name) {
    // You would think that the below does the trick...
    XSSFTable table = sheet.createTable(ref);
    // ..but it produces a corrupt table with all the IDs wrong, so they need fixing as below...
    for (int i = 1; i < table.getCTTable().getTableColumns().sizeOfTableColumnArray(); i++) {
      table.getCTTable().getTableColumns().getTableColumnArray(i).setId(i + 1);
    }
    table.setName(name);
    table.setDisplayName(name);
    // table.getCTTable().setId(10);
    // For now, create the initial style in a low-level way
    table.getCTTable().addNewTableStyleInfo();
    table.getCTTable().getTableStyleInfo().setName("TableStyleMedium2");

    // Style the table
    XSSFTableStyleInfo style = (XSSFTableStyleInfo) table.getStyle();
    style.setName("TableStyleMedium2");
    style.setShowColumnStripes(false);
    style.setShowRowStripes(true);
    style.setFirstColumn(false);
    style.setLastColumn(false);
    style.setShowRowStripes(true);
    // style.setShowColumnStripes(true);
    /*
            // Set the values for the table
            XSSFRow row;
            XSSFCell cell;
            for (int i = 0; i < 3; i++) {
                // Create row
                row = sheet.createRow(i);
                for (int j = 0; j < 3; j++) {
                    // Create cell
                    cell = row.createCell(j);
                    if (i == 0) {
                        cell.setCellValue("Column" + (j + 1));
                    } else {
                        cell.setCellValue((i + 1.0) * (j + 1.0));
                    }
                }
            }
    */
    table.getCTTable().addNewAutoFilter().setRef(table.getArea().formatAsString());
    return table;
  }

  /** Force get a cell (create if it does not exist) */
  public static XSSFCell findAndCreateCellIfRequired(
      Workbook workbook, String sheetName, int row, int column, JSONArray potentialErrors) {
    Sheet s = workbook.getSheet(sheetName);
    if (s == null) {
      potentialErrors.put("Sheet [" + sheetName + "] was not found");
      return null;
    }
    Row r = s.getRow(row);
    if (r == null) r = s.createRow(row);
    Cell c = r.getCell(column);
    if (c == null) c = r.createCell(column);
    return (XSSFCell) c;
  }

  /*
   * Insert JSON data into cell, taking type conversions into account
   */

  public static boolean insertDataFromJSONIntoCell(Object value, XSSFCell c) {
    return insertDataFromJSONIntoCell(value, c, null, null);
  }

  public static boolean insertDataFromJSONIntoCell(
      Object value, XSSFCell c, String format, Integer indent) {
    if (c == null) throw new IllegalArgumentException("Cell was null");
    if (value == null) {
      c.setBlank();
      // c.setCellValue("");
    }
    // This is the unwrapping step - in case the value provided is JSON, then
    // 1) Take "data" as the value
    // 2) Check for "format" for cell-level formatting
    // 3) Check for "indent" for cell-level indentation
    else if (value instanceof JSONObject) {
      JSONObject json = (JSONObject) value;
      // System.out.println("Cool: " + json.toString());
      if (json.has("data")) {
        format = json.has("format") ? json.getString("format") : null;
        indent =
            json.has("indent") && json.get("indent") instanceof Integer
                ? json.getInt("indent")
                : null;
        Object data = json.get("data");
        insertDataFromJSONIntoCell(data, c, format, indent);
      }
    } else if (value instanceof Boolean) {
      c.setCellValue((boolean) value);
    } else if (value instanceof Number) {
      c.setCellValue(((Number) value).doubleValue());
    } else {
      if (Utils.isFormattedAsDate(value.toString())) {
        c.setCellValue(Utils.getAsDate(value.toString()));
      } else if (value.toString().startsWith("=")) {
        c.setCellFormula(value.toString().substring(1));
      } else {
        c.setCellValue(value.toString());
      }
    }
    // Set cell-level formatting
    if (format != null || indent != null) {
      XSSFCellStyle cellStyle = c.getSheet().getWorkbook().createCellStyle();
      // Make sure we start from the existing cell's style
      cellStyle.cloneStyleFrom(c.getCellStyle());
      if (format != null) {
        // Swap "." and "," for locale-specific number formats
        if (format.contains(".") && format.contains(","))
          format = format.replace(".", "DOT").replace(",", ".").replace("DOT", ",");
        XSSFDataFormat dataFormat = c.getSheet().getWorkbook().createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(format));
      }
      if (indent != null) {
        short s = indent.shortValue();
        cellStyle.setIndention(s);
        cellStyle.setAlignment(HorizontalAlignment.LEFT); // Discovered by Peter Albert
        log.info("Setting indent [" + s + "] for cell [" + c.getAddress().formatAsString() + "]");
      }
      c.setCellStyle(cellStyle);
      if (format != null) return true;
    }
    return false;
  }
}
