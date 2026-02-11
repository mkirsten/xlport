package com.molnify.xlport.core;

import com.molnify.xlport.pdf.ExportFormat;
import com.molnify.xlport.pdf.PDFExporter;
import org.junit.Ignore;
import org.junit.Test;

public class PDFExporterTest {

  @Ignore("Requires Google credentials")
  @Test
  public void testToLink() throws Exception {
    Template template1 = TemplateManager.getTemplate("template1.xlsx");
    System.out.println(PDFExporter.uploadAndReturnId(template1.workbook));
    template1.workbook.close();
  }

  @Ignore("Requires Google credentials")
  @Test
  public void testToFile() throws Exception {
    Template template1 = TemplateManager.getTemplate("template1.xlsx");
    System.out.println(PDFExporter.toFile(template1.workbook, new ExportFormat()));
    template1.workbook.close();
  }
}
