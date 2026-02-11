package com.molnify.xlport.util;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class PowerQueryBug {

  // TODO: Move this to other tests (and make it a real test, not just a stand alone class)
  private static final String FILE_NAME =
      "src/test/resources/test-suites/export/query_table_simple/template copy.xlsx";

  public static void main(String args[]) {
    try {
      FileInputStream inputStream = new FileInputStream(new File(FILE_NAME));
      Workbook workbook = WorkbookFactory.create(inputStream);

      Sheet sheet = workbook.getSheetAt(0);
      sheet.createRow(2);
      Row r = sheet.getRow(2);
      r.createCell(0);
      r.createCell(1);
      r.createCell(2);

      r.getCell(0).setCellValue(10);
      r.getCell(1).setCellValue("Option 1");
      r.getCell(2).setCellValue("ck9ubsezi49h70913on5nqehs");

      inputStream.close();

      FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
      evaluator.evaluateAll();

      FileOutputStream outputStream = new FileOutputStream(FILE_NAME.replace("copy.", "updated."));
      workbook.write(outputStream);
      workbook.close();
      outputStream.close();

    } catch (IOException | EncryptedDocumentException ex) {
      ex.printStackTrace();
    }
  }
}
