package org.example.excel_io.utils;

import com.google.api.services.sheets.v4.Sheets;
import org.example.excel_io.api.Credentials;

public class ExportGoogleToGoogle {
    private final Sheets service;
    private final String spreadsheetId = "1Zf_9P0Ewy2N-fTfmQSkJLkckBAB62rko7zXDV6mTh0E"; // тестовый ID

    /**
     * Проходим авторизацию API Google Sheets в момент создания объекта
     */
    public ExportGoogleToGoogle() {
        service = Credentials.getSheets();
    }

}
