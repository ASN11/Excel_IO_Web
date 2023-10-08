package org.example.excel_io.views;

import com.vaadin.flow.component.UI;
import com.vaadin.flow.component.button.Button;
import com.vaadin.flow.component.dialog.Dialog;
import com.vaadin.flow.component.html.H3;
import com.vaadin.flow.component.html.Span;
import com.vaadin.flow.component.orderedlayout.HorizontalLayout;
import com.vaadin.flow.component.orderedlayout.VerticalLayout;
import com.vaadin.flow.component.page.Page;
import com.vaadin.flow.component.upload.Upload;
import com.vaadin.flow.component.upload.receivers.MemoryBuffer;
import com.vaadin.flow.router.Route;
import org.example.excel_io.utils.ExcelReader;
import org.example.excel_io.utils.ExportExcelToGoogle;
import org.springframework.beans.factory.annotation.Autowired;

import java.io.*;
import java.util.List;

import static java.io.File.separator;


@Route("/")
public class MainFormView extends VerticalLayout {
    private final Upload singleFileUpload;
    private final Button saveButton = new Button("START");
    private FileInputStream fileInputStream;
    private String fileName;

@Autowired
    public MainFormView(MemoryBuffer memoryBuffer) {
        this.singleFileUpload = new Upload(memoryBuffer);

        getFileInputStreamFromExcel(memoryBuffer);
        buttonStart();
        add(new HorizontalLayout(singleFileUpload, saveButton));
    }

    /**
     * Создает поток fileInputStream на основе файла, который выбрал пользователь в singleFileUpload
     */
    private void getFileInputStreamFromExcel(MemoryBuffer memoryBuffer) {
        singleFileUpload.addSucceededListener(event -> {
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
                e.printStackTrace();
            }
        });
    }

    private void buttonStart() {
        saveButton.addClickListener(click -> {
            ExcelReader excelReader = new ExcelReader(fileInputStream);
            excelReader.analyze("result_" + fileName);

            ExportExcelToGoogle exportExcelToGoogle = new ExportExcelToGoogle();
            exportExcelToGoogle.export("src" + separator + "main" + separator + "resources" + separator + "result_" + fileName);

            excelReader.delete("result_" + fileName);

            if (exportExcelToGoogle.courierListIsEmpty()) {
                createSuccessMessage();
            } else {
                createErrorMessage(exportExcelToGoogle.getCourierList());
            }
        });
    }

    /**
     * Создает всплывающее окно, уведомляюзее об успешном заполнении промежуточной таблицы и кнопкой заполнения итоговой таблицы
     */
    private void createSuccessMessage() {
        Dialog successDialog = new Dialog();
        successDialog.setWidth("600px");

        H3 successHeader = new H3("Промежуточная таблицa заполнена успешно");
        successHeader.getStyle().set("color", "green");

        Button successButton = new Button("Заполнить итоговую таблицу", event -> {
            successDialog.close();
            exportGoogleToGoogle();
        });

        VerticalLayout successLayout = new VerticalLayout(successHeader, successButton);
        successLayout.setAlignItems(Alignment.CENTER);
        successLayout.setSpacing(true);

        successDialog.add(successLayout);
        successDialog.open();
    }

    /**
     * Перенос данных из промежуточной таблицы в итоговую
     */
    private void exportGoogleToGoogle() {

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


