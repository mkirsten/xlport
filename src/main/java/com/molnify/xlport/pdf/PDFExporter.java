package com.molnify.xlport.pdf;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.FileContent;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.api.services.drive.DriveScopes;
import com.google.api.services.drive.model.Permission;
import com.molnify.xlport.core.Utils;
import com.molnify.xlport.servlet.InitXlPort;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Collections;
import java.util.List;
import java.util.Random;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * This is a class that manages export (from Excel) to PDF
 *
 * @author kirsten
 */
public class PDFExporter {

  private static final Logger log = Logger.getLogger(PDFExporter.class.getName());

  private static GoogleCredential credential = null;
  private static final List<String> SCOPES = Collections.singletonList(DriveScopes.DRIVE_FILE);

  public static String uploadAndReturnId(Workbook workbook) throws Exception {
    String tmpFileName = "xlport-temp-pdfexport" + new Random().nextInt(1000000),
        tmpSuffix = ".xlsx";
    File copiedFile = File.createTempFile(tmpFileName, tmpSuffix);
    try (FileOutputStream fileOutputStream = new FileOutputStream(copiedFile); ) {
      workbook.write(fileOutputStream);
    } catch (IOException e) {
      throw e;
    }
    String MIME = "application/vnd.google-apps.spreadsheet";
    Permission permission = new Permission().setType("anyone").setRole("writer");
    com.google.api.services.drive.model.File fileMetaData =
        new com.google.api.services.drive.model.File();
    fileMetaData.setName(tmpFileName + tmpSuffix);
    fileMetaData.setMimeType(MIME);
    FileContent fc = new FileContent(MIME, copiedFile);
    com.google.api.services.drive.model.File uploadedFile =
        driveApi().files().create(fileMetaData, fc).setFields("id").execute();
    String id = uploadedFile.getId();
    driveApi().permissions().create(id, permission).execute();
    return id;
  }

  public static String toFile(Workbook workbook, ExportFormat format) throws Exception {
    String id = uploadAndReturnId(workbook);
    String url = format.getExportURLForId(id);
    File tmp = Utils.saveContentsOfUrlAsTmpFile(url);
    driveApi().files().delete(id);
    return tmp.getAbsolutePath();
  }

  private static Drive driveApi() {
    JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();
    HttpTransport transport = new NetHttpTransport();
    return new Drive.Builder(transport, jsonFactory, credential)
        .setApplicationName("xlport")
        .build();
  }

  static {
    try {
      InputStream fis =
          new ByteArrayInputStream(
              InitXlPort.getGoogleCredential().toString().getBytes(StandardCharsets.UTF_8));
      credential = GoogleCredential.fromStream(fis).createScoped(SCOPES);
      log.info(
          "Loaded Google credential (for sheets access) for service account: "
              + credential.getServiceAccountId());
    } catch (IOException e) {
      log.info("Error loading credentials data.");
    }
  }
}
