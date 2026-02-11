package com.molnify.xlport.core;

import java.io.Closeable;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Container for an Excel template workbook and its metadata.
 *
 * <p>Tracks template items (named ranges and tables), sheet-scoped items for multi-sheet
 * templates, and provides workbook protection. Implements {@link Closeable} to ensure
 * the underlying workbook is properly closed.
 */
public class Template implements Closeable {
  public String originalFileName = null;
  public HashMap<String, TemplateItem> items = new HashMap<>();
  // This is here to allow for sheet scoped template items. When properties have been copied as a
  // part of the logic
  // to create new sheets from a template sheet, the named ranges become sheet scoped
  public HashMap<String, Map<String, TemplateItem>> sheetScopedItems = new HashMap<>();
  public XSSFWorkbook workbook = null;

  public void addTemplateItem(TemplateItem item) {
    items.put(item.name, item);
  }

  public void addTemplateItemScoped(TemplateItem item, String sheet) {
    if (sheetScopedItems.get(sheet) == null)
      sheetScopedItems.put(sheet, new HashMap<String, TemplateItem>());
    sheetScopedItems.get(sheet).put(item.name, item);
  }

  public String getOriginalFileSuffix() {
    if (originalFileName == null) return "";
    else {
      int dotPosition = originalFileName.indexOf(".");
      if (dotPosition < 0 || dotPosition == originalFileName.length()) return "";
      else return originalFileName.substring(dotPosition + 1);
    }
  }

  public void protectWorkbook(String password) {
    if (workbook != null) {
      for (Sheet s : workbook) {
        XSSFSheet sheet = ((XSSFSheet) s);

        if (password != null && password.length() > 0) sheet.protectSheet(password);

        sheet.enableLocking();
        sheet.lockInsertRows(true);
        sheet.lockInsertColumns(true);
        sheet.lockDeleteRows(true);
        sheet.lockDeleteColumns(true);
        sheet.lockObjects(true);
        sheet.lockSelectLockedCells(false);
        sheet.lockSelectUnlockedCells(false);
      }
    }
  }

  @Override
  public String toString() {
    StringBuilder buf = new StringBuilder();
    buf.append("TEMPLATE\n");
    for (TemplateItem item : items.values()) buf.append(item.toString()).append("\n");
    for (Map.Entry<String, Map<String, TemplateItem>> e : sheetScopedItems.entrySet())
      buf.append("Sheet [" + e.getKey() + "]: " + e.getValue().toString()).append("\n");

    return buf.toString();
  }

  @Override
  public void close() throws IOException {
    if (workbook != null) workbook.close();
  }
}
