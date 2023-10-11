package org.example.excel_io.views;

import com.vaadin.flow.component.UI;
import com.vaadin.flow.component.button.Button;
import com.vaadin.flow.component.dialog.Dialog;
import com.vaadin.flow.component.html.H3;
import com.vaadin.flow.component.html.Span;
import com.vaadin.flow.component.notification.Notification;
import com.vaadin.flow.component.notification.NotificationVariant;
import com.vaadin.flow.component.orderedlayout.HorizontalLayout;
import com.vaadin.flow.component.orderedlayout.VerticalLayout;
import com.vaadin.flow.component.page.Page;
import com.vaadin.flow.component.upload.Upload;
import com.vaadin.flow.component.upload.receivers.MemoryBuffer;
import com.vaadin.flow.router.Route;
import org.example.excel_io.utils.ExcelReader;
import org.example.excel_io.utils.ExportExcelToGoogle;
import org.example.excel_io.utils.ExportGoogleToGoogle;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.*;
import java.util.List;
import java.text.SimpleDateFormat;
import java.util.Date;

import static java.io.File.separator;


@Route("/")
public class MainFormView extends VerticalLayout {
    private final Upload singleFileUpload;
    private final Button firstButton = new Button("START");
    private final Button secondButton = new Button("Заполнить промежуточную таблицу");
    private FileInputStream fileInputStream;
    private String fileName;
    private ExcelReader excelReader;

@Autowired
    public MainFormView(MemoryBuffer memoryBuffer) {
        this.singleFileUpload = new Upload(memoryBuffer);
        secondButton.setVisible(false);

        getFileInputStreamFromExcel(memoryBuffer);
        firstButton();
        secondButton();
        add(new HorizontalLayout(singleFileUpload, firstButton, secondButton));
    }

    /**
     * Создает поток fileInputStream на основе файла, который выбрал пользователь в singleFileUpload
     */
    private void getFileInputStreamFromExcel(MemoryBuffer memoryBuffer) {
        singleFileUpload.addSucceededListener(event -> {
            secondButton.setVisible(false);
            try {
                fileName = event.getFileName();
                byte[] fileBytes = memoryBuffer.getInputStream().readAllBytes();

                // Создаем временный файл и записываем в него данные
                File tempFile = File.createTempFile("uploaded-", fileName);
                try (FileOutputStream fileOutputStream = new FileOutputStream(tempFile)) {
                    fileOutputStream.write(fileBytes);
                }

                // Теперь у вас есть FileInputStream для временного файла
                fileInputStream = new FileInputStream(tempFile);

            } catch (IOException e) {
                createNotification(e.getMessage(), NotificationVariant.LUMO_ERROR);
            }
        });
    }

    /**
     * Обработка кнопки "START"
     */
    private void firstButton() {
        firstButton.addClickListener(click -> {
            excelReader = new ExcelReader(fileInputStream);
            try {
                excelReader.analyze("result_" + fileName);
                createNotification("Файл успешно обработан", NotificationVariant.LUMO_SUCCESS);
                secondButton.setVisible(true);
            } catch (Exception e) {
                createNotification(e.getMessage(), NotificationVariant.LUMO_ERROR);
            }
        });
    }

    /**
     * Обработка кнопки "Заполнить промежуточную таблицу"
     */
    private void secondButton() {
        secondButton.addClickListener(click -> {
            try {
                ExportExcelToGoogle exportExcelToGoogle = new ExportExcelToGoogle();
                exportExcelToGoogle.export("src" + separator + "main" + separator + "resources" + separator + "result_" + fileName);

                excelReader.delete("result_" + fileName);

                if (exportExcelToGoogle.courierListIsEmpty()) {
                    createSuccessMessage("Промежуточная таблицa заполнена успешно",
                            "Заполнить итоговую таблицу",
                            "https://docs.google.com/spreadsheets/d/1Zf_9P0Ewy2N-fTfmQSkJLkckBAB62rko7zXDV6mTh0E/edit#gid=527366521");
                } else {
                    createErrorMessage(exportExcelToGoogle.getCourierList());
                }
            } catch (Exception e) {
                createNotification(e.getMessage(), NotificationVariant.LUMO_ERROR);
            }
        });
    }

    /**
     * Отображает всплывающее уведомление об ошибке
     */
    private void createNotification(String message, NotificationVariant variant) {
        Notification errorNotification = new Notification( message, 5000, Notification.Position.BOTTOM_CENTER);
        errorNotification.addThemeVariants(variant);

        errorNotification.open();
    }

    /**
     * Создает всплывающее окно, уведомляюзее об успешном заполнении таблицы и кнопкой заполнения итоговой таблицы (или закрытия)
     */
    private void createSuccessMessage(String header, String successButtonText, String url) {
        Dialog successDialog = new Dialog();
        successDialog.setWidth("600px");

        H3 successHeader = new H3(header);
        successHeader.getStyle().set("color", "green");

        VerticalLayout successLayout = getSuccessLayout(successDialog, successHeader, successButtonText, url);
        successLayout.setAlignItems(Alignment.CENTER);
        successLayout.setSpacing(true);

        successDialog.add(successLayout);
        successDialog.open();
    }

    private VerticalLayout getSuccessLayout(Dialog successDialog, H3 successHeader, String successButtonText, String url) {
        Button successButton = new Button(successButtonText, event -> {
            successDialog.close();
            if (successButtonText.equals("Заполнить итоговую таблицу")) {
                exportGoogleToGoogle();
            }
        });

        Button openTable = new Button("Открыть таблицу", event -> {
            Page page = new Page(UI.getCurrent());
            page.open(url, "_blank");
        });

        return new VerticalLayout(successHeader, new HorizontalLayout(openTable, successButton));
    }

    /**
     * Перенос данных из промежуточной таблицы в итоговую
     */
    private void exportGoogleToGoogle() {
        ExportGoogleToGoogle exportGoogleToGoogle = new ExportGoogleToGoogle();
        try {
            int destinationSheetID = exportGoogleToGoogle.copy(today());
            createSuccessMessage("Итоговая таблицa заполнена успешно",
                    "Закрыть",
                    "https://docs.google.com/spreadsheets/d/1BxBh2BYRj8Kb38E54QPnOiwYJ5ymwQVuSMsBOMKO6g4/edit#gid=" + destinationSheetID);
        } catch (IOException e) {
            createNotification(e.getMessage(), NotificationVariant.LUMO_ERROR);
        }
    }

    /**
     * Создает строку в фортате "dd.MM.yyyy" с текущей датой
     */
    private String today() {
        // Получаем текущую дату
        Date currentDate = new Date();

        // Создаем объект SimpleDateFormat для форматирования
        SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");

        // Форматируем дату и выводим её
        return dateFormat.format(currentDate);
    }

    /**
     * Создает всплывающее окно, в котором отображает список курьеров, которых нет в справочнике
     */
    private void createErrorMessage(List<String> courierList) {
        Dialog errorDialog = new Dialog();
        errorDialog.setWidth("600px");

        H3 errorHeader = new H3("Нет данных на курьеров:");
        errorHeader.getStyle().set("color", "red");
        VerticalLayout errorLayout = getVerticalLayout(courierList, errorHeader, errorDialog);
        errorLayout.setAlignItems(Alignment.CENTER);
        errorLayout.setSpacing(true);

        errorDialog.add(errorLayout);
        errorDialog.open();
    }

    private static VerticalLayout getVerticalLayout(List<String> courierList, H3 errorHeader, Dialog errorDialog) {
        VerticalLayout errorLayout = new VerticalLayout(errorHeader);

        for (String line : courierList) {
            Span span = new Span(line);
            errorLayout.add(span);
        }

        Button closeButton = new Button("Закрыть", event -> errorDialog.close());
        Button errorPlaceButton = new Button("Заполнить данные", event -> {
            Page page = new Page(UI.getCurrent());
            page.open("https://docs.google.com/spreadsheets/d/1Zf_9P0Ewy2N-fTfmQSkJLkckBAB62rko7zXDV6mTh0E/edit#gid=527366521",
                    "_blank");
        });

        HorizontalLayout horizontalLayout = new HorizontalLayout(errorPlaceButton, closeButton);

        errorLayout.add(horizontalLayout);
        return errorLayout;
    }
}


