package com.molnify.xlport.core;

import static org.junit.Assert.*;
import static org.junit.Assume.assumeTrue;

import com.molnify.xlport.servlet.InitXlPort;
import java.io.IOException;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class WorkbookFetcherGoogleSheetsTest {

  private final String url =
      "https://docs.google.com/spreadsheets/d/1kI796WUrKoYg-pe_eFIj5iWL10QU8KpYm7uhyjElQpU/edit#gid=0";

  @Test
  public void testGetId() {
    String id = WorkbookFetcherGoogleSheets.getIdFromUrl(url);
    assertEquals("1kI796WUrKoYg-pe_eFIj5iWL10QU8KpYm7uhyjElQpU", id);
  }

  @Test
  public void testFetch() throws IOException, ClassNotFoundException {
    assumeTrue("Google credentials not configured", InitXlPort.GOOGLE_CREDENTIAL != null);
    XSSFWorkbook wb = WorkbookFetcherGoogleSheets.fetchGoogleSheetsTemplate(url);
    assertNotNull(wb);
    String value = wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue();
    assertEquals("hejsan", value);
  }
}
