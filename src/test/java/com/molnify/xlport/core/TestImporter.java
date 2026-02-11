package com.molnify.xlport.core;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertTrue;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;

public class TestImporter {
  @Test
  public void testBasicImport() {
    // Test data
    Map<String, String> testProps = new HashMap<String, String>();
    testProps.put("LastUpdated", "2018-12-14T14:51:12.977Z");
    testProps.put("ProjectId", "uuid2");
    testProps.put("ProjectLabel", "Project name");

    // Prep request and execute import
    Template template = TemplateManager.getTemplate("template1.xlsx");
    Workbook workbook = template.workbook;
    JSONObject request = new JSONObject().put(Importer.PROPERTIES, testProps.keySet());
    JSONArray potentialErrors = new JSONArray();
    JSONObject result =
        Importer.importData(request, (XSSFWorkbook) workbook, potentialErrors, false);

    // Validate result
    assertEquals(0, potentialErrors.length());
    assertTrue(result.has(Importer.PROPERTIES));
    JSONObject props = result.getJSONObject(Importer.PROPERTIES);
    assertEquals(testProps.keySet().size(), props.keySet().size());
    for (String key : testProps.keySet()) assertEquals(testProps.get(key), props.getString(key));
  }

  @Test
  public void testDataTypes() throws EncryptedDocumentException, IOException {
    // Test data
    JSONObject expected = new JSONObject();
    expected.put("Blank", JSONObject.NULL);
    expected.put("String", "Great stuff");
    expected.put("Integer", 42.0);
    expected.put("Double", 123.789d);
    expected.put("Error", "#NAME?"); // <-- Only require this starts with Error
    expected.put("Date", "1982-01-25T00:00:00.000Z");
    expected.put("Formula_Int", 43.0);
    expected.put("Formula_Date", "2019-01-20T00:00:00.000Z");
    expected.put("Formula_String", "First second");
    expected.put("Formula_Error", "#DIV/0!"); // <-- Only require this starts with Error

    // Prep request and execute import
    Template template = TemplateManager.getLocalTemplateInTestDirectory("import-datatypes.xlsx");
    Workbook workbook = template.workbook;
    JSONObject request = new JSONObject().put(Importer.PROPERTIES, expected.keySet());
    JSONArray potentialErrors = new JSONArray();
    JSONObject result =
        Importer.importData(request, (XSSFWorkbook) workbook, potentialErrors, false);

    // Validate result
    assertEquals(0, potentialErrors.length());
    assertTrue(result.has(Importer.PROPERTIES));
    JSONObject props = result.getJSONObject(Importer.PROPERTIES);
    assertEquals(expected.keySet().size(), props.keySet().size());
    for (String key : expected.keySet()) {
      if (expected.get(key) instanceof String && expected.getString(key).startsWith("#"))
        assertTrue(props.getString(key).startsWith("#"));
      else assertEquals(expected.get(key), props.get(key));
    }
  }

  @Test
  public void testTableImport() {
    // Test data
    List<String> tablesToImport = new ArrayList<String>();
    tablesToImport.add("Initiatives");
    tablesToImport.add("Status");
    tablesToImport.add("Table4");

    // Prep request and execute import
    Template template = TemplateManager.getTemplate("template1.xlsx");
    Workbook workbook = template.workbook;
    JSONObject request = new JSONObject().put(Importer.TABLES, tablesToImport);
    JSONArray potentialErrors = new JSONArray();
    JSONObject result =
        Importer.importData(request, (XSSFWorkbook) workbook, potentialErrors, false);

    // Validate result
    assertEquals(0, potentialErrors.length());
    assertTrue(result.has(Importer.TABLES));
    JSONObject tables = result.getJSONObject(Importer.TABLES);
    assertEquals(tablesToImport.size(), tables.keySet().size());
  }

  @Test
  public void testWildcards() {
    // Prep request and execute import
    Template template = TemplateManager.getTemplate("template1.xlsx");
    Workbook workbook = template.workbook;
    JSONObject request =
        new JSONObject()
            .put(Importer.TABLES, new JSONArray().put("*"))
            .put(Importer.PROPERTIES, new JSONArray().put("*"));
    JSONArray potentialErrors = new JSONArray();
    JSONObject result =
        Importer.importData(request, (XSSFWorkbook) workbook, potentialErrors, false);

    // Validate result
    assertEquals(0, potentialErrors.length());
    assertTrue(result.has(Importer.PROPERTIES));
    assertEquals(3, result.getJSONObject(Importer.PROPERTIES).keySet().size());
    assertTrue(result.has(Importer.TABLES));
    assertEquals(3, result.getJSONObject(Importer.TABLES).keySet().size());
  }

  @Test
  public void extractMultipleValuesFromSingleCellInTable() throws IOException {
    Workbook wb =
        WorkbookFactory.create(
            new File("src/test/resources/extract-multiple-values-from-single-cell-in-table.xlsx"));
    JSONArray potentialErrors = new JSONArray();
    JSONObject result = Importer.importAllData(wb, potentialErrors, true);
    assertEquals(0, potentialErrors.length());
    assertNotNull(result);
  }
}
