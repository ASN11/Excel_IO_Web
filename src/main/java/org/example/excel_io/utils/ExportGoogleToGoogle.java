package org.example.excel_io.utils;

import com.google.api.services.sheets.v4.Sheets;
import com.google.api.services.sheets.v4.model.*;
import lombok.Getter;
import org.example.excel_io.api.Credentials;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

@Getter
@Component
public class ExportGoogleToGoogle {
    private final Sheets service;

    @Value("${table.medium.id}")
    private String mediumSpreadsheetId;
    @Value("${table.result.id}")
    private String resultSpreadsheetId;
    @Value("${source_sheet.id}")
    private int sourceSheetID;

    /**
     * Проходим авторизацию API Google Sheets в момент создания объекта
     */
    @Autowired
    public ExportGoogleToGoogle(Credentials credentials) {
        service = credentials.getSheets();
    }

    public int copy(String newSheetName) throws IOException {
        int destinationSheetID = createNewSheet(newSheetName);
        copyData(newSheetName);

        return destinationSheetID;
    }

    /**
     * Копирует данные из локальной в итоговую таблицу
     */
    private void copyData(String newSheetName) throws IOException {
        String RANGE_1_OLD = "Основной файл!A2:C500";
        String RANGE_1_NEW = newSheetName + "!A5:C503";
        String RANGE_2_OLD = "Основной файл!E2:H500";
        String RANGE_2_NEW = newSheetName + "!D5:G503";
        String RANGE_3_OLD = "Основной файл!D2:D500";
        String RANGE_3_NEW = newSheetName + "!H5:H503";

        copyRange(RANGE_1_OLD, RANGE_1_NEW);
        copyRange(RANGE_2_OLD, RANGE_2_NEW);
        copyRange(RANGE_3_OLD, RANGE_3_NEW);
    }

    private void copyRange(String RANGE_OLD, String RANGE_NEW) throws IOException {
        ValueRange response1 = service.spreadsheets().values().get(mediumSpreadsheetId, RANGE_OLD).execute();
        List<List<Object>> values1 = response1.getValues();
        ValueRange body1 = new ValueRange().setRange(RANGE_NEW).setValues(values1);

        // Экспорт значений в Google Таблицы
        service.spreadsheets().values()
                .update(resultSpreadsheetId, RANGE_NEW, body1).setValueInputOption("USER_ENTERED")
                .execute();
    }

    /**
     * Копирует лист ID 58633884 и создаёт на его основе лист newSheetName
     */
    private int createNewSheet(String newSheetName) throws IOException {
        // Копируем лист со всеми свойствами
        DuplicateSheetRequest duplicateRequest = new DuplicateSheetRequest()
                .setSourceSheetId(sourceSheetID)
                .setNewSheetName(newSheetName);

        List<Request> requests = new ArrayList<>();
        requests.add(new Request()
                .setDuplicateSheet(duplicateRequest));

        // Выполняем запрос на копирование листа
        BatchUpdateSpreadsheetRequest body = new BatchUpdateSpreadsheetRequest().setRequests(requests);
        BatchUpdateSpreadsheetResponse response = service.spreadsheets().batchUpdate(resultSpreadsheetId, body).execute();

        // Извлекаем SheetId созданного листа
        SheetProperties createdSheetProperties = response.getReplies().get(0).getDuplicateSheet().getProperties();

        return createdSheetProperties.getSheetId();
    }

}
