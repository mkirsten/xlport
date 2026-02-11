package com.molnify.xlport.core;

import static org.junit.Assert.*;

import java.io.File;
import java.io.IOException;
import static org.junit.Assume.assumeTrue;

import java.util.Arrays;
import java.util.List;
import java.util.function.Consumer;
import java.util.logging.Logger;
import com.molnify.xlport.servlet.InitXlPort;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeUtil;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Ignore;
import org.junit.Test;

/**
 * This runs all test suites in the folder for the export tests Key idea is that it should be easy
 * to add new tests. A suite consists of Template - the template for export Request - a JSON file
 * with data and a reference to the template to use Expected - the expected Excel file that should
 * be produced by the service
 *
 * @author kirsten
 */
public class TestSuiteExporter {
  private static final String REQUEST = "request.json",
      EXPECTED = "expected.xlsx",
      TEMPLATE = "template.xlsx",
      TEST_DIRECTORY = "src/test/resources/test-suites/export/";

  private static final Logger log = Logger.getLogger(TestSuiteExporter.class.getName());

  private static final String[] knownTestSuites =
      new String[] {
        "1object",
        "1table",
        "1table-10col-100rows",
        "multipleObjects",
        "multipleObjectsAndTables",
        "multipleTables",
        "1table-with-calculations",
        "formulasAndLookups",
        "deals",
        "daysbug",
        "sheets",
        "query_table_simple",
        "query_table_bug",
        "formula_not_overwritten",
        "thedate",
        "firstXlPort2Test",
        "secondXlPort2Test",
        "thirdXlPort2Test",
        "v2-multi-sheet-column",
        "v2-multi-column-with-fixed-column",
        "v2-multi-sheet-simple",
        "v2-multi-sheet-column-simple",
        "v2-multi-sheet-column-fixed",
        "v2-alma-scorecard-export",
        "v2-utc",
        "v2-multi-sheet-column-fixed-topleft",
        "v2-alma-minimized",
        "v2-alma-table",
        "v2-column-formatting",
        "v2-multi-sheet-column-simple-dash",
        "v2-formats",
        "lookup",
        "dataValidation",
        "v2-1object",
        "v2-almascorecard-24",
        "v2-multi-sheet-bug",
        "v2-table-lookup",
        "1table-10col-100rows-withnull"
      };

  @Test
  public void test1object() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[0]));
  }

  @Test
  public void test1table() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[1]));
  }

  @Test
  public void test1table10col100rows() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[2]));
  }

  @Test
  public void testMultipleObjects() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[3]));
  }

  @Test
  public void testMultipleObjectsAndTables() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[4]));
  }

  @Test
  public void testMultipleTables() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[5]));
  }

  @Test
  public void test1tableWithCalculations() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[6]));
  }

  @Test
  public void testFormulasAndLookups() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[7]));
  }

  @Ignore("Disabled - needs investigation")
  @Test
  public void testDeals() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[8]));
  }

  @Test
  public void testDaysbug() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[9]));
  }

  @Test
  public void testGoogleSheets() throws Exception {
    assumeTrue("Google credentials not configured", InitXlPort.GOOGLE_CREDENTIAL != null);
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[10]));
  }

  @Test
  public void testQueryTableSimple() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[11]));
  }

  @Test
  public void testQueryTableBug() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[12]));
  }

  @Test
  public void testFormulaNotOverwritten() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[13]));
  }

  @Test
  public void testTheDate() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[14]));
  }

  @Test
  public void testFirstXlPort2Test() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[15]));
  }

  @Test
  public void testSecondXlPort2Test() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[16]));
  }

  @Test
  public void testThirdXlPort2Test() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[17]));
  }

  @Test
  public void testV2MultiSheetColumn() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[18]));
  }

  @Test
  public void testV2MultiColumnWithFixedColumn() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[19]));
  }

  @Test
  public void testV2MultiSheetSimple() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[20]));
  }

  @Test
  public void testV2MultiSheetColumnSimple() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[21]));
  }

  @Test
  public void testV2MultiSheetColumnFixed() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[22]));
  }

  @Ignore("Disabled - needs investigation")
  @Test
  public void testV2AlmaScorecardExport() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[23]));
  }

  @Test
  public void testV2utc() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[24]));
  }

  @Test
  public void testV2MultiSheetColumnFixedTopleft() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[25]));
  }

  @Test
  public void testV2AlmaScorecardMinimized() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[26]));
  }

  @Test
  public void testV2AlmaTable() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[27]));
  }

  @Test
  public void testV2ColumnFormatting() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[28]));
  }

  @Test
  public void testV2MultiSheetColumnSimpleDash() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[29]));
  }

  @Test
  public void test1table10col100rowsWithNull() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[37]));
  }

  @Ignore("Format handling needs rework")
  @Test
  public void testV2Formats() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[30]));
  }

  @Ignore("Broken since Apache POI 5.2.0")
  @Test
  public void testLookup() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[31]));
  }

  @Test
  public void testDataValidation() throws Exception {
    Consumer<Workbook[]> additionalTest =
        (Workbook[] workbooks) -> {
          Workbook expected = workbooks[0], actual = workbooks[1];
          DataValidation eDv = expected.getSheet("Data").getDataValidations().get(0),
              aDv = actual.getSheet("Data").getDataValidations().get(0);
          StringBuilder eList = new StringBuilder(), aList = new StringBuilder();
          Arrays.asList(eDv.getRegions().getCellRangeAddresses())
              .forEach((CellRangeAddress c) -> eList.append(c.formatAsString()).append("   "));
          Arrays.asList(aDv.getRegions().getCellRangeAddresses())
              .forEach((CellRangeAddress c) -> aList.append(c.formatAsString()).append("   "));

          // If expected is e.g., A1:A7 (a single range), make a special check for convenience
          if (eDv.getRegions().getCellRangeAddresses().length == 1) {
            CellRangeAddress ex = eDv.getRegions().getCellRangeAddress(0);
            int numberChecked = 0;
            for (CellRangeAddress c : aDv.getRegions().getCellRangeAddresses()) {
              assertTrue(
                  "The regions for data validations appears to be incorrect. "
                      + ex.formatAsString()
                      + " is in expected but here is a cell in the actual outside that range: "
                      + c.formatAsString(),
                  CellRangeUtil.contains(ex, c));
              numberChecked++;
            }
            assertEquals(
                "Didn't find enough actual data validation cells to check", 2, numberChecked);
          } else
            assertEquals(
                "Data validation in the first sheet should be for the same ranges, but they are"
                    + " not",
                eList.toString().trim(),
                aList.toString().trim());
        };
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[32]), additionalTest);
  }

  @Test
  public void testV2_1object() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[33]));
  }

  @Ignore("Not yet passing")
  @Test
  public void testV2almascorecard24() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[34]));
  }

  @Test
  public void testV2multiSheetBug() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[35]));
  }

  @Test
  public void testV2tableLookup() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[36]));
  }

  @Ignore("Discovery test - run manually")
  @Test
  public void testAllOtherSuitesForExport() throws Exception {
    File[] directoryForTests = new File(TEST_DIRECTORY).listFiles();
    if (directoryForTests == null) fail("Could not find test directory: " + TEST_DIRECTORY);
    Arrays.sort(directoryForTests);
    for (File dir : directoryForTests) {
      if (Arrays.asList(knownTestSuites).contains(dir.getName())) continue;
      log.info("Found unlisted test that will also be run [" + dir.getName() + "]");
      runSingleTestInSuite(dir);
    }
  }

  /**
   * Run a test from a directory. The files expected are template.xlsx - The template file
   * request.json - The request JSON expected.xlsx - The expected xlsx file from using the template
   * and applying data from the request file
   *
   * @param dir The directory
   * @throws EncryptedDocumentException
   * @throws IOException
   */
  private static void runSingleTestInSuite(File dir)
      throws EncryptedDocumentException, IOException {
    runSingleTestInSuite(dir, null);
  }

  /**
   * Same as the method without the additionalTestMethod, apart from
   *
   * @param dir The directory
   * @param additionalTestMethod A method (Workbook[] -> Boolean) to also run for specific tests to
   *     check. The array contains two workbooks, the expected and the actual
   * @throws EncryptedDocumentException
   * @throws IOException
   */
  private static void runSingleTestInSuite(File dir, Consumer<Workbook[]> additionalTestMethod)
      throws EncryptedDocumentException, IOException {
    if (!dir.isDirectory()) return;
    String path = dir.getAbsolutePath() + "/";
    log.info("Running test [" + dir.getName() + "] in path [" + path + "]");

    // Set up the 3 files that are in each folder
    JSONObject json = new JSONObject(Utils.readFileAsString(path + REQUEST, true));
    File expectedFile = new File(path + EXPECTED);
    Workbook expected = WorkbookFactory.create(expectedFile, null, true);
    Template template;
    if (dir.getName().equals("sheets"))
      template = TemplateManager.getTemplate(json.getString("templateId"));
    else template = TemplateManager.getLocalTemplateFromDirectory(path, TEMPLATE);

    if (json.keySet().size() == 0) fail("no request provided");

    // Export without and reported errors
    try {
      JSONArray potentialErrors = new JSONArray();
      Exporter.exportToExcel(json.getJSONObject("data"), template, potentialErrors, true);
      if (potentialErrors.length() > 0)
        fail(
            "Test in folder [" + dir.getCanonicalPath() + "] contained errors: " + potentialErrors);

      // Ensure the created workbook matched the expected
      File out = File.createTempFile("exported", ".xlsx");
      Utils.writeOutWorkbookAsFile(template, out);
      log.info("Exported file can be found @ " + out.getAbsolutePath());
      log.info("Expected file can be found @ " + path + EXPECTED);

      List<String> errors = Utils.diffTwoWorkbooksAndReturnErrors(expected, template.workbook);
      if (errors.size() != 0) {
        StringBuilder s = new StringBuilder();
        for (String e : errors) s.append("\n").append(e);
        fail("Error in test [" + dir.getName() + "]: " + s);
      }

      if (additionalTestMethod != null)
        additionalTestMethod.accept(new Workbook[] {expected, template.workbook});
    } catch (Exception e) {
      e.printStackTrace();
      fail(e.getMessage());
    } finally {
      if (template != null) template.close();
      expected.close();
    }
    log.info("Test [" + dir.getName() + "] passed");
  }
}
