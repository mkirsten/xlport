package com.molnify.xlport.core;

import static org.junit.Assert.*;

import com.molnify.xlport.TestUtils;
import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Test;

public class UtilsTest {

  @Test
  public void testRemoveAndInsertColumns() throws Exception {
    XSSFWorkbook workbook =
        TestUtils.createWorkbookFromResource("deleteAndInsertColumns/testToInsertColumn.xlsx");
    Utils.copyColumn("Sheet1", 3, workbook.getSheet("Sheet2"), 5);
    Utils.removeColumn(workbook.getSheet("Sheet1"), 3);
    Utils.removeCalcChain(workbook);
    // Requires evaluation, as this is based on =TODAY()
    XSSFFormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
    formulaEvaluator.evaluateAll();

    XSSFWorkbook expectedWorkbook =
        TestUtils.createWorkbookFromResource(
            "deleteAndInsertColumns/testToInsertColumn-updated.xlsx");
    List<String> errors = Utils.diffTwoWorkbooksAndReturnErrors(expectedWorkbook, workbook);

    assertEquals((errors.size() > 0) ? errors.get(0) : "no error", 0, errors.size());
  }

  @Test
  public void expandNormalName() {
    String[] names = Utils.expandName("normal");
    assertEquals(1, names.length);
    assertEquals("normal", names[0]);
    String[] names2 = Utils.expandName("normal_withunderscore");
    assertEquals(1, names2.length);
    assertEquals("normal_withunderscore", names2[0]);
  }

  @Test
  public void collapseAndExpandName() {
    List<String[]> namesToTest = new ArrayList<>();
    namesToTest.add(new String[] {"SheetName", "ItemName"});
    namesToTest.add(new String[] {"Sheet Name", "ItemName"});
    namesToTest.add(new String[] {"Sheet_Name", "ItemName"});

    for (String[] s : namesToTest) {
      assertEquals(
          Arrays.toString(s), Arrays.toString(Utils.expandName(Utils.collapseName(s[0], s[1]))));
      // System.out.println(Arrays.toString(s) + " : " + collapsed + " : " +
      // Arrays.toString(expanded));
    }
  }

  @Test
  public void testIsFormattedAsDate() {
    assertFalse(Utils.isFormattedAsDate("no"));
    assertFalse(Utils.isFormattedAsDate("1982-01-25"));
    assertFalse(Utils.isFormattedAsDate("1982-01-25:04:50"));
    assertTrue(Utils.isFormattedAsDate("2019-01-25T14:53:12.977Z"));
    assertTrue(Utils.isFormattedAsDate("2019-01-25T14:53:12.97Z"));
  }

  @Test
  public void testGetAsDate() {
    assertNotNull(Utils.getAsDate("2019-01-25T14:53:12.977Z"));
    assertNotNull(Utils.getAsDate("1982-01-25T04:53:10.12Z"));
    assertNull(Utils.getAsDate("1982-01-25"));
  }

  @Test
  public void testDiffTwoWorkbooksAndReturnFirstError() throws IOException {
    Workbook original = TemplateManager.getTemplate("dummy").workbook;
    Workbook actual = TemplateManager.getTemplate("dummy").workbook;

    assertEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, original).size());
    assertEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, actual).size());

    final int SHEET = 2, ROW = 3, COL = 1;
    actual.getSheetAt(SHEET).getRow(ROW).getCell(COL).setCellValue("new!");
    assertNotEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, actual).size());
    String oldValue = original.getSheetAt(SHEET).getRow(ROW).getCell(COL).getStringCellValue();
    actual.getSheetAt(SHEET).getRow(ROW).getCell(COL).setCellValue(oldValue);
    assertEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, actual).size());

    actual.createSheet("new sheet");
    assertNotEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, actual).size());
    actual.removeSheetAt(actual.getNumberOfSheets() - 1);
    assertEquals(0, Utils.diffTwoWorkbooksAndReturnErrors(original, actual).size());

    original.close();
    actual.close();
  }
}
