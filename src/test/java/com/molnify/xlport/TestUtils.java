package com.molnify.xlport;

import java.io.InputStream;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/*
 * Class for utility functions used in tests
 */
public class TestUtils {

  /**
   * Create a XSSFWorkbook object from a given resource path.
   *
   * @param resourcePath The relative path to the Excel workbook resource. (e.g. "test.xlsx" from
   *     /src/test/resources/test.xlsx")
   * @return An XSSFWorkbook object representing the loaded Excel workbook.
   * @throws IllegalArgumentException If the resource cannot be found at the specified path.
   */
  public static XSSFWorkbook createWorkbookFromResource(String resourcePath) throws Exception {
    InputStream inputStream = TestUtils.class.getClassLoader().getResourceAsStream(resourcePath);
    if (inputStream == null) {
      throw new IllegalArgumentException("Resource not found: " + resourcePath);
    }
    return (XSSFWorkbook) WorkbookFactory.create(inputStream);
  }
}
