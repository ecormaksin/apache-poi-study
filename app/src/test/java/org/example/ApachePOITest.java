package org.example;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.xwpf.usermodel.BodyType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.*;
import java.nio.file.Files;

public class ApachePOITest {

    @ParameterizedTest
    @ValueSource(strings = {"template_without_table.doc", "template_with_table.doc"})
    void testHWPF(String templateFileName) throws Exception {
        File templateFile = new File("src/main/resources/" + templateFileName);
        File outputFile = new File("/tmp/" + templateFileName);

        try (InputStream inputStream = Files.newInputStream(templateFile.toPath());
             HWPFDocument aHWPFDocument = new HWPFDocument(inputStream);) {

            try (OutputStream outputStream = Files.newOutputStream(outputFile.toPath())) {
                aHWPFDocument.write(outputStream);
            }
        }

        try( InputStream inputStream = Files.newInputStream(outputFile.toPath());
        HWPFDocument aHWPFDocument = new HWPFDocument(inputStream);) {
            Range aRange = aHWPFDocument.getRange();
        }
    }

    @ParameterizedTest
    @ValueSource(strings = {"template_without_table.docx", "template_with_table.docx"})
    void testXWPF(String templateFileName) throws Exception {
        File templateFile = new File("src/main/resources/" + templateFileName);
        File outputFile = new File("/tmp/" + templateFileName);

        try (InputStream inputStream = Files.newInputStream(templateFile.toPath());
             XWPFDocument aXWPFDocument = new XWPFDocument(inputStream);) {

            try (OutputStream outputStream = Files.newOutputStream(outputFile.toPath())) {
                aXWPFDocument.write(outputStream);
            }
        }

        try( InputStream inputStream = Files.newInputStream(outputFile.toPath());
             XWPFDocument aHWPFDocument = new XWPFDocument(inputStream);) {
            BodyType aBodyType = aHWPFDocument.getPartType();
        }
    }
}
