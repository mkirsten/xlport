package com.molnify.xlport.core;

import static org.junit.Assert.fail;

import com.google.common.collect.Maps;
import com.google.common.reflect.TypeToken;
import com.google.gson.Gson;
import java.io.File;
import java.io.IOException;
import java.lang.reflect.Type;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.util.Map;
import java.util.logging.Logger;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;
import org.json.JSONArray;
import org.json.JSONObject;
import org.junit.Test;

/**
 * Some tests for import
 *
 * @author kirsten
 */
public class TestSuiteImporter {
  private static final String REQUEST = "request.json",
      EXPECTED = "expected.json",
      WORKBOOK = "workbook.xlsx",
      TEST_DIRECTORY = "src/test/resources/test-suites/import/";

  private static final Logger log = Logger.getLogger(TestSuiteImporter.class.getName());

  private static final String[] knownTestSuites =
      new String[] {"1object", "1table", "1table-10col-100rows", "v2importMultipleSheets"};

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
  public void testV2importMultipleSheets() throws Exception {
    runSingleTestInSuite(new File(TEST_DIRECTORY + knownTestSuites[3]));
  }

  private static void runSingleTestInSuite(File dir)
      throws EncryptedDocumentException, IOException {
    if (!dir.isDirectory()) return;
    String path = dir.getAbsolutePath() + "/";
    log.info("Running test [" + dir.getName() + "] in path [" + path + "]");

    // Set up the 3 files that are in each folder
    JSONObject request = null;
    try {
      request = new JSONObject(Utils.readFileAsString(path + REQUEST, true));
    } catch (Exception e) {
      log.info("No request.json provided for this test. Using default request");
    }
    JSONObject expected = new JSONObject(Utils.readFileAsString(path + EXPECTED, true));
    if (expected == null) fail("no expected file provided");
    File workbookFile = new File(path + WORKBOOK);
    XSSFWorkbook workbook = new XSSFWorkbookFactory().create(workbookFile, null, true);

    try {
      JSONArray potentialErrors = new JSONArray();
      JSONObject result;
      if (request == null) result = Importer.importAllData(workbook, potentialErrors, true);
      else result = Importer.importData(request, workbook, potentialErrors, true);

      if (potentialErrors.length() > 0)
        fail(
            "Test in folder [" + dir.getCanonicalPath() + "] contained errors: " + potentialErrors);
      File out = File.createTempFile("result", ".json");
      Files.write(out.toPath(), result.toString().getBytes(StandardCharsets.UTF_8));
      log.info("Exported file can be found @ " + out.getAbsolutePath());
      log.info("Expected file can be found @ " + path + EXPECTED);

      // Ensure the created workbook matched the expected
      Gson g = new Gson();
      Type mapType = new TypeToken<Map<String, Object>>() {}.getType();
      Map<String, Object> firstMap = g.fromJson(expected.toString(), mapType);
      Map<String, Object> secondMap = g.fromJson(result.toString(), mapType);
      String errorMessage = Maps.difference(firstMap, secondMap).toString();
      if (!"equal".equals(errorMessage))
        fail("Error in test [" + dir.getName() + "]: " + errorMessage);
    } catch (Exception e) {
      e.printStackTrace();
      fail(e.getMessage());
    } finally {
      if (workbook != null) workbook.close();
    }
    log.info("Test [" + dir.getName() + "] passed");
  }
}
