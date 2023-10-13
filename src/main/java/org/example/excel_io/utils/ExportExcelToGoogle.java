package org.example.excel_io.utils;

import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.aspose.cells.WorksheetCollection;
import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.ClearValuesRequest;
import com.google.api.services.sheets.v4.model.Spreadsheet;
import com.google.api.services.sheets.v4.model.UpdateValuesResponse;
import com.google.api.services.sheets.v4.model.ValueRange;
import org.example.excel_io.api.Credentials;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Component
public class ExportExcelToGoogle {
    private final Sheets service;

    @Value("${table.medium.id}")
    private String spreadsheetId;

    /**
     * Проходим авторизацию API Google Sheets в момент создания объекта
     */
    @Autowired
    public ExportExcelToGoogle(Credentials credentials) {
        service = credentials.getSheets();
    }

    public void export(String _excelFileName) throws Exception {
        // Подключение к существующей таблице по её ID
        Spreadsheet spreadsheet = service.spreadsheets().get(spreadsheetId).execute();

        if (new File(_excelFileName).exists()) {
            // Получить ссылку на первый рабочий лист
            Worksheet ws = createFirstWorksheet(_excelFileName);
            List<List<Object>> worksheetData = getDataFromExcel(ws);

            String range = "Основной файл!A2:E500";
            clearRequest(service, range);
            UpdateValuesResponse result = setDataToGoogleSheet(range, worksheetData, spreadsheet);

            if (result.getUpdatedCells() == 0) {
                throw new IOException("Некорректный файл, данные не записаны");
            }

            getCourierList();
        }
    }

    /**
     * Проверка на наличие данных на курьера в справочнике
     */
    public List<String> getCourierList() throws IOException {
        String range = "Основной файл!E2:F500";
        List<String> courierList = new ArrayList<>();

        ValueRange response = service.spreadsheets().values()
                .get(spreadsheetId, range)
                .execute();

        List<List<Object>> values = response.getValues();
        if (values != null && !values.isEmpty()) {
            for (List<Object> row : values) {
                if (!row.get(0).equals("") && row.get(1).equals(" ")) {
                    courierList.add((String) row.get(0)); // Значение в столбце E
                }
            }
        }

        return courierList;
    }

    public boolean courierListIsEmpty() throws IOException {
        List<String> courierList = getCourierList();

        return courierList.isEmpty();
    }

    /**
     * Записываем данные из worksheetData в Google таблицу
     */
    private UpdateValuesResponse setDataToGoogleSheet(String range, List<List<Object>> worksheetData, Spreadsheet spreadsheet) throws IOException {
        // Установить диапазон
        ValueRange body = new ValueRange();
        body.setRange(range);

        // Установить значения
        body.setValues(worksheetData);

        // Экспорт значений в Google Таблицы
        UpdateValuesResponse result = service.spreadsheets().values()
                .update(spreadsheet.getSpreadsheetId(), range, body).setValueInputOption("USER_ENTERED")
                .execute();

        return result;
    }

    /**
     * Переписывает все данные с листа Excel ws в лист worksheetData
     */
    private List<List<Object>> getDataFromExcel(Worksheet ws) {
        // Получить количество строк и столбцов
        int rows = ws.getCells().getMaxDataRow();
        int cols = ws.getCells().getMaxDataColumn();

        List<List<Object>> worksheetData = new ArrayList<>();

        // Цикл по строкам
        for (int i = 0; i <= rows; i++) {
            List<Object> rowData = new ArrayList<>();

            for (int j = 0; j <= cols; j++) {
                if (ws.getCells().get(i, j).getValue() == null)
                    rowData.add("");
                else
                    rowData.add(ws.getCells().get(i, j).getValue());
            }

            worksheetData.add(rowData);
        }

        return worksheetData;
    }

    /**
     * Возвращает первый рабочий лист из файла _excelFileName
     */
    private Worksheet createFirstWorksheet(String _excelFileName) throws Exception {
        // Загрузите книгу Excel
        Workbook wb = new Workbook(_excelFileName);
        // Получить все рабочие листы
        WorksheetCollection collection = wb.getWorksheets();

        return collection.get(0);
    }

    /**
     * Выполняет запрос на очистку диапазона range
     */
    private void clearRequest(Sheets service, String range) throws IOException {
        ClearValuesRequest clearRequest = new ClearValuesRequest();
        service.spreadsheets().values().clear(spreadsheetId, range, clearRequest).execute();
    }
}
