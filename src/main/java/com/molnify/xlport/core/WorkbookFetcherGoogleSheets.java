package com.molnify.xlport.core;

import com.google.api.client.googleapis.auth.oauth2.GoogleCredential;
import com.google.api.client.http.HttpTransport;
import com.google.api.client.http.javanet.NetHttpTransport;
import com.google.api.client.json.JsonFactory;
import com.google.api.client.json.jackson2.JacksonFactory;
import com.google.api.services.drive.Drive;
import com.google.common.collect.ImmutableList;
import com.molnify.xlport.servlet.InitXlPort;
import java.io.ByteArrayInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * This provides functionality to use a Google Sheet as a Template
 *
 * @author kirsten
 */
public class WorkbookFetcherGoogleSheets {

  protected static XSSFWorkbook fetchGoogleSheetsTemplate(String url) throws IOException {
    if (InitXlPort.GOOGLE_CREDENTIAL == null)
      throw new ExceptionInInitializerError(
          "Credentials not set up. You need to provide credentials as part of setup (if you use"
              + " xlport as a library) or set environment variables. Check the class InitXlPort for"
              + " more details");
    XSSFWorkbook x = null;
    InputStream fis =
        new ByteArrayInputStream(InitXlPort.GOOGLE_CREDENTIAL.getBytes(StandardCharsets.UTF_8));
    GoogleCredential credential = GoogleCredential.fromStream(fis);
    credential =
        credential.createScoped(
            ImmutableList.of(
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"));

    JsonFactory jsonFactory = JacksonFactory.getDefaultInstance();
    HttpTransport transport = new NetHttpTransport();
    Drive build =
        new Drive.Builder(transport, jsonFactory, credential).setApplicationName("xlport").build();

    try {
      String id = getIdFromUrl(url);
      InputStream is =
          build
              .files()
              .export(id, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
              .executeAsInputStream();
      x = new XSSFWorkbook(is);
      is.close();
    } catch (IOException e) {
      throw e;
    }
    return x;
  }

  protected static String getIdFromUrl(String url) {
    if (url == null) return null;
    String[] items = url.split("/");
    String longest = "";
    for (String s : items) {
      if (s.length() > longest.length()) longest = s;
    }
    return longest;
  }
}
