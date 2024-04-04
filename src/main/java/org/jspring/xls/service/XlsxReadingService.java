package org.jspring.xls.service;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayInputStream;
import java.io.FileInputStream;
import java.io.IOException;


/**
 * The XlsxReadingService class provides functionality for reading data from an XLSX file.
 */
public class XlsxReadingService {

    private final String templatePath;
    public XlsxReadingService(String templatePath) {
        this.templatePath = templatePath;
    }

    /**
     * Reads an XSSFWorkbook from a template file.
     *
     * @return The XSSFWorkbook read from the template file.
     * @throws RuntimeException if an IOException occurs while reading the template file.
     */
    public XSSFWorkbook readFromTemplate() {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Reads an {@link XSSFWorkbook} from a template file.
     *
     * @param templatePath The path of the template file.
     * @return The {@link XSSFWorkbook} read from the template file.
     * @throws RuntimeException if an {@link IOException} occurs while reading the template file.
     */
    public XSSFWorkbook readFromTemplate(String templatePath) {
        try (FileInputStream inputStream = new FileInputStream(templatePath)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * Reads an XSSFWorkbook from a byte array.
     *
     * @param fileContent The byte array containing the workbook data.
     * @return The XSSFWorkbook read from the byte array.
     * @throws RuntimeException if an IOException occurs while reading the byte array.
     */
    public XSSFWorkbook readFromByteArray(byte[] fileContent) {
        try (ByteArrayInputStream inputStream = new ByteArrayInputStream(fileContent)) {
            return new XSSFWorkbook(inputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

}
