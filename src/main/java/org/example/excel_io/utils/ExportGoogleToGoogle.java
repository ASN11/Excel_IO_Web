package org.example.excel_io.utils;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import org.example.excel_io.api.Credentials;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class ExportGoogleToGoogle {
    private final Sheets service;
    private final String mediumSpreadsheetId = "1Zf_9P0Ewy2N-fTfmQSkJLkckBAB62rko7zXDV6mTh0E"; // тестовый ID
    private final String resultSpreadsheetId = "1BxBh2BYRj8Kb38E54QPnOiwYJ5ymwQVuSMsBOMKO6g4"; // тестовый ID

    /**
     * Проходим авторизацию API Google Sheets в момент создания объекта
     */
    public ExportGoogleToGoogle() {
        service = Credentials.getSheets();
    }

    public void copy() {
        copySheet();


    }

    private void copySheet() {
        try {
            // Имя исходного и целевого листов
            String sourceSheetName = "Шаблон (не трогать)";
            int sourceSheetID = 58633884;
            String destinationSheetName = "Тестовый лист_1";
            int destinationSheetID = 600103475;

            // Копируем ширину столбцов из исходного листа в целевой
            DuplicateSheetRequest duplicateRequest = new DuplicateSheetRequest()
                    .setSourceSheetId(sourceSheetID)
                    .setNewSheetName(destinationSheetName);

            List<Request> requests = new ArrayList<>();
            requests.add(new Request()
                    .setDuplicateSheet(duplicateRequest));

            // Выполняем запрос на копирование листа
            BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
            service.spreadsheets().batchUpdate(resultSpreadsheetId, body).execute();

        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    private Integer getSheetId(String destinationSheetName) throws IOException {
        // После добавления нового листа, выполните запрос на получение информации о таблице
        Spreadsheet spreadsheet = service.spreadsheets().get(resultSpreadsheetId).execute();

        // Получите список листов в таблице
        List<Sheet> sheets = spreadsheet.getSheets();

        // Ищите лист с именем "Тестовый лист_1" и получите его sheetId
        Integer sheetId = null;
        for (Sheet sheet : sheets) {
            if (sheet.getProperties().getTitle().equals(destinationSheetName)) {
                sheetId = sheet.getProperties().getSheetId();
                break;
            }
        }

        if (sheetId != null) {
            System.out.println("sheetId нового листа 'Тестовый лист_1': " + sheetId);
        } else {
            System.out.println("Лист 'Тестовый лист_1' не найден в таблице.");
        }

        return sheetId;
    }


}
