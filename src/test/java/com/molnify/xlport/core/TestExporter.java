package com.molnify.xlport.core;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.fail;

import java.io.IOException;
import org.json.JSONArray;
import org.json.JSONException;
import org.json.JSONObject;
import org.junit.Test;

public class TestExporter {

  @Test
  public void test() throws IOException {
    assertCertainNumberOfErrorsInExport("src/test/resources/export1.json", "dummy", 0);
    assertCertainNumberOfErrorsInExport("src/test/resources/export2.json", "dummy", 0);
    assertCertainNumberOfErrorsInExport("src/test/resources/export3.json", "dummy", 0);

    try {
      assertCertainNumberOfErrorsInExport("src/test/resources/invalid-export1.json", "dummy", 1);
      fail("Should fail on JSON");
    } catch (JSONException e) {
      // As expected from the test
    }
    assertCertainNumberOfErrorsInExport("src/test/resources/invalid-export2.json", "dummy", 0);
    try {
      assertCertainNumberOfErrorsInExport("src/test/resources/invalid-export3.json", "dummy", 1);
      fail("Should fail on JSON");
    } catch (JSONException e) {
      // As expected from the test
    }
  }

  private void assertCertainNumberOfErrorsInExport(
      String jsonRequestFileName, String templateId, int numberOfErrors) throws IOException {
    String exportRequest = Utils.readFileAsString(jsonRequestFileName, false);
    JSONObject json = new JSONObject(exportRequest);
    Template template = TemplateManager.getTemplate(templateId);
    JSONArray potentialErrors = new JSONArray();
    Exporter.exportToExcel(json.getJSONObject("data"), template, potentialErrors, true);
    assertEquals(numberOfErrors, potentialErrors.length());
    template.close();
  }
}
