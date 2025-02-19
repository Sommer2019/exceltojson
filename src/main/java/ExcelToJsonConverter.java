import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelToJsonConverter {

    private static final String SETTINGS_SHEET = "Einstellungen";
    private static final List<String> EXCLUDED_SHEETS = Arrays.asList("Vorlage");

    public static void main(String[] args) {
        String filePath = "src/main/resources/Mappe1.xlsm";
        String outputDir = "out/tests";

        Map<String, Map<String, String>> columnMappings = new HashMap<>();
        Map<String, List<Map<String, String>>> excelData = new HashMap<>();

        try (FileInputStream fis = new FileInputStream(new File(filePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet settingsSheet = workbook.getSheet(SETTINGS_SHEET);
            if (settingsSheet != null) {
                columnMappings = loadSettingsMappings(settingsSheet);
            }

            for (Sheet sheet : workbook) {
                String sheetName = sheet.getSheetName();

                if (EXCLUDED_SHEETS.contains(sheetName) || SETTINGS_SHEET.equals(sheetName)) {
                    continue;
                }

                if (isB1Filled(sheet)) {
                    continue;
                }

                excelData.put(sheetName, readSheetWithMappings(sheet, columnMappings));
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        saveJsonFiles(excelData, outputDir);
    }

    public static Map<String, Map<String, String>> loadSettingsMappings(Sheet sheet) {
        Map<String, Map<String, String>> mappings = new HashMap<>();
        List<String> headers = new ArrayList<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                for (int col = 1; col < row.getLastCellNum(); col++) {
                    Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    headers.add(cell.toString().trim());
                }
            } else {
                Cell firstCell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                String category = firstCell.toString().trim();

                if (!category.isEmpty()) {
                    Map<String, String> columnMapping = new HashMap<>();
                    for (int col = 1; col < headers.size() + 1; col++) {
                        Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        String mappedValue = cell.toString().trim();
                        if (!mappedValue.isEmpty()) {
                            columnMapping.put(headers.get(col - 1), mappedValue);
                        }
                    }
                    mappings.put(category, columnMapping);
                }
            }
        }
        return mappings;
    }

    public static boolean isB1Filled(Sheet sheet) {
        Row firstRow = sheet.getRow(0);
        if (firstRow != null) {
            Cell b1Cell = firstRow.getCell(1);
            return b1Cell != null && !b1Cell.toString().trim().isEmpty();
        }
        return false;
    }

    public static List<Map<String, String>> readSheetWithMappings(Sheet sheet, Map<String, Map<String, String>> columnMappings) {
        List<Map<String, String>> sheetData = new ArrayList<>();
        List<String> headers = new ArrayList<>();

        for (Row row : sheet) {
            if (row.getRowNum() == 0) {
                continue;
            }

            if (row.getRowNum() == 1) {
                for (int col = 1; col < row.getLastCellNum(); col++) {
                    Cell cell = row.getCell(col, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    headers.add(cell.toString().trim());
                }
                continue;
            }

            Cell categoryCell = row.getCell(0, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
            String category = categoryCell.toString().trim();

            if (!category.isEmpty() && columnMappings.containsKey(category)) {
                Map<String, String> rowData = new LinkedHashMap<>();
                Map<String, String> mapping = columnMappings.get(category);

                for (int i = 0; i < headers.size(); i++) {
                    Cell cell = row.getCell(i+1, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                    String originalHeader = headers.get(i);
                    if (mapping.containsKey(originalHeader)) {
                        String jsonField = mapping.get(originalHeader);
                        String value = cell.toString().trim();
                        rowData.put(jsonField, value);
                    }
                }
                sheetData.add(rowData);
            }
        }
        return sheetData;
    }

    public static void saveJsonFiles(Map<String, List<Map<String, String>>> excelData, String outputDir) {
        ObjectMapper objectMapper = new ObjectMapper();
        new File(outputDir).mkdirs();

        for (Map.Entry<String, List<Map<String, String>>> entry : excelData.entrySet()) {
            String sheetName = entry.getKey();
            List<Map<String, String>> data = entry.getValue();

            File outputFile = new File(outputDir, sheetName + ".json");
            try {
                objectMapper.writeValue(outputFile, data);
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}
