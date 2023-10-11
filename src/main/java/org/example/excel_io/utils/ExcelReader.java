package org.example.excel_io.utils;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static java.io.File.separator;

public class ExcelReader {
    private final FileInputStream sourceFileInputStream;

    public ExcelReader(FileInputStream sourceFileInputStream) {
        this.sourceFileInputStream = sourceFileInputStream;
    }

    public void delete(String fileName) {
        String targetFilePath = "src" + separator + "main" + separator + "resources" + separator + fileName;
        File fileToDelete = new File(targetFilePath);

        if (fileToDelete.delete()) {
            System.out.println("Файл удален успешно.");
        } else {
            System.err.println("Не удалось удалить файл.");
        }
    }
    public void analyze(String fileName) throws IOException {
        // Путь к целевому Excel-файлу
        String targetFilePath = "src" + separator + "main" + separator + "resources" + separator + fileName;

        // Создаем рабочую книгу (workbook) из исходного файла
        Workbook sourceWorkbook = new XSSFWorkbook(sourceFileInputStream);

        // Создаем новую книгу для записи данных
        Workbook targetWorkbook = new XSSFWorkbook();

        // Получаем листы из исходной и целевой книг
        Sheet sourceSheet = sourceWorkbook.getSheetAt(0);
        Sheet targetSheet = targetWorkbook.createSheet("Новый лист");

        List<String> columnDataA = extractColumnData(0, sourceSheet);
        List<String> columnDataB = extractColumnData(1, sourceSheet);
        List<String> columnDataC = extractColumnData(2, sourceSheet);
        Map<String, String> names = getNames(sourceSheet);

        copyData(sourceSheet, targetSheet, columnDataA, columnDataB, columnDataC, names);

        // Создаем FileOutputStream для записи целевого файла
        FileOutputStream targetFileOutputStream = new FileOutputStream(targetFilePath);

        // Сохраняем целевую книгу в целевом файле
        targetWorkbook.write(targetFileOutputStream);

        targetFileOutputStream.close();
    }

    /**
     *  Копируем данные из исходного листа в целевой лист
     */
    private void copyData(Sheet sourceSheet, Sheet targetSheet, List<String> columnDataA, List<String> columnDataB, List<String> columnDataC, Map<String, String> names) {
        int rowResultIndex = 0;

        for (int rowIndex = 0; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row sourceRow = sourceSheet.getRow(rowIndex);
            if (sourceRow != null) {

                // удаляем маршруты, где вообще нет Строгино
                if (rowIndex > 0 && !columnDataA.get(rowIndex).equals(columnDataA.get(rowIndex-1)) && columnDataA.get(rowIndex).contains("Маршрут")) {
                    int rowIndexLocal = rowIndex;
                    boolean flag = true;
                    do {
                        if (columnDataC.get(rowIndexLocal).equals("Москва 2-я лыковская 63с6")) {
                            flag = false;
                            break;
                        }
                        rowIndexLocal++;
                    }
                    while (rowIndexLocal != columnDataA.size() && columnDataA.get(rowIndexLocal).equals(columnDataA.get(rowIndexLocal-1)));

                    if (flag) {
                        rowIndex = rowIndexLocal;
                        continue;
                    }
                }

                // удаляем строку Shop&Show
                if(sourceRow.getCell(2).toString().contains("Shop&Show")) {
                    continue;
                }

                // удаляем часть маршрута,которые начинаются на Тарном, а заканчиваются на Строгино
                if (rowIndex > 0 && !columnDataA.get(rowIndex).equals(columnDataA.get(rowIndex-1)) && columnDataA.get(rowIndex).contains("Маршрут")) {
                    int rowIndexLocal = rowIndex;
                    boolean flag = false;

                    do {
                        if (columnDataC.get(rowIndexLocal).equals("Москва, Промышленная, д. 12А")) {
                            flag = true;
                            break;
                        }
                        rowIndexLocal++;
                    }
                    while (rowIndexLocal != columnDataA.size() && columnDataA.get(rowIndexLocal).equals(columnDataA.get(rowIndexLocal-1)) &&
                            !columnDataC.get(rowIndexLocal).equals("Москва 2-я лыковская 63с6"));

                    if (flag) {
                        rowIndex = rowIndexLocal;
                        continue;
                    }
                }

                // если это возврат, то пропускаем строку
                if(sourceRow.getCell(1).toString().equals("Доставка")) {
                    continue;
                }

                // если это забор возврата с СЦ, то пропускаем строку
                if(sourceRow.getCell(2).toString().startsWith("Тарный ") ||
                        sourceRow.getCell(2).toString().startsWith("Софьино ") ||
                        sourceRow.getCell(2).toString().startsWith("Томилино ")) {
                    continue;
                }

                // Удаляем все, кроме маршрутов и пробелов между ними
                if(!sourceRow.getCell(0).toString().contains("Маршрут") && !sourceRow.getCell(0).toString().isEmpty()) {
                    continue;
                }

                // Удаляем дублирование строки с СЦ
                if (columnDataC.get(rowIndex).equals("Москва 2-я лыковская 63с6") && columnDataC.get(rowIndex).equals(columnDataC.get(rowIndex-1))) {
                    continue;
                }

                // удаляем часть маршрута,которые начинаются на Строгино, а заканчиваются на Тарном
                if (columnDataB.get(rowIndex).equals("Склад")  && columnDataA.get(rowIndex).equals(columnDataA.get(rowIndex-1))) {
                    int rowIndexLocal = rowIndex;
                    while (!columnDataC.get(rowIndexLocal).isEmpty()) {
                        if (columnDataC.get(rowIndexLocal).equals("Москва, Промышленная, д. 12А") || rowIndexLocal == columnDataC.size()-1) {
                            rowIndex = rowIndexLocal;
                            break;
                        }
                        rowIndexLocal++;
                    }
                }


                Row resultRow = targetSheet.createRow(rowResultIndex);

                int cellResultIndex = 0;
                for (int cellIndex = 0; cellIndex < sourceRow.getLastCellNum(); cellIndex++) {

                    // Удаляем столбец со старыми именами
                    if (cellIndex == 5) {
                        continue;
                    }

                    Cell sourceCell = sourceRow.getCell(cellIndex);
                    Cell resultCell = resultRow.createCell(cellResultIndex);

                    // Чтобы не создавать новые ячейки, подменяем столбец со временем отправения
                    if(cellResultIndex == 4) {
                        if (names.containsKey(sourceRow.getCell(0).toString())) {
                            sourceCell.setCellValue(names.get(sourceRow.getCell(0).toString()));
                            names.remove(sourceRow.getCell(0).toString());
                        } else {
                            sourceCell.setCellValue("");
                        }
                    }

                    // удаляем лишнее время забора
                    if ((!sourceRow.getCell(2).toString().equals("Москва 2-я лыковская 63с6") || columnDataC.get(rowIndex-1).isEmpty()
                            || columnDataC.get(rowIndex-1).equals("Номер заказа")) && cellIndex == 3) {
                        sourceCell.setCellValue("");
                    }

                    rewriteSourceToResult(sourceCell, resultCell);

                    cellResultIndex++;
                }
            }
            rowResultIndex++;
        }
    }

    /**
     * Заполняет вспомогательные массивы из исходной таблицы для анализа данных
     *
     * @param i номер столбца в таблице
     * @param sourceSheet лист, из которого берем данные
     */
    private  List<String> extractColumnData(int i, Sheet sourceSheet) {
        List<String> columnData = new ArrayList<>();

        for (int rowIndex = 0; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row row = sourceSheet.getRow(rowIndex);

            if (row != null) {
                Cell cell = row.getCell(i);

                if (cell != null) {
                    columnData.add(cell.toString());
                } else {
                    columnData.add("");
                }
            } else {
                columnData.add("");
            }
        }
        return columnData;
    }

    /**
     * Проходим по таблице и собираем всех курьеров с их маршрутами
     *
     * @param sourceSheet лист, из которого берем данные
     * @return Map (ключ - маршрут курьера, значение - имя курьера)
     */
    private Map<String, String> getNames(Sheet sourceSheet) {
        Map<String, String> names = new HashMap<>();

        for (int rowIndex = 0; rowIndex <= sourceSheet.getLastRowNum(); rowIndex++) {
            Row row = sourceSheet.getRow(rowIndex);

            if (row != null) {
                Cell cell = row.getCell(5);

                if (cell != null) {
                    names.put(row.getCell(0).toString(), cell.toString());
                }
            }
        }

         return names;
    }

    /**
     * Копирует данные из исходной ячейки в итоговую
     *
     * @param sourceCell исходная ячейка
     * @param resultCell итоговая ячейка
     */
    private void rewriteSourceToResult(Cell sourceCell, Cell resultCell) {
        if (sourceCell != null) {
            switch (sourceCell.getCellType()) {
                case STRING:
                    resultCell.setCellValue(sourceCell.getStringCellValue());
                    break;
                case NUMERIC:
                    resultCell.setCellValue(sourceCell.getNumericCellValue());
                    break;
                case BOOLEAN:
                    resultCell.setCellValue(sourceCell.getBooleanCellValue());
                    break;
                case FORMULA:
                    resultCell.setCellFormula(sourceCell.getCellFormula());
                    break;
                default:
                    resultCell.setCellValue("");
            }
        }
    }
}
