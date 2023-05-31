package org.example.Util;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class GeneradorExcel {

    public static void exportExcel(String pestana, String titulo, int registros, List<String> listaCabecera, List<List<Object>> listaAtributos) throws FileNotFoundException {
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss dd/MM/yyyy");
        SimpleDateFormat sdfHora = new SimpleDateFormat("HH:mm:ss");
        SimpleDateFormat sdfFecha = new SimpleDateFormat("dd/MM/yyyy");
        final int NUMERO_CELDAS = 107;
        int numeroCabeceras = listaCabecera.size();

        try {
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            Sheet sheet = workbook.createSheet(pestana);

            // en el codigo del BO no será necesario
            InputStream is = new FileInputStream("C:\\data\\logo-ripley.png");
            byte[] logoReporte = IOUtils.toByteArray(is);
            byte[] celeste = new byte[]{(byte) 0, (byte) 255, (byte) 255};
            XSSFColor colorCabecera = new XSSFColor(celeste);

            // FONDO BLANCO
            //fillColorInCells(sheet, listaAtributos);

            //FUENTE
            // TITULO
            Font fontTitle = workbook.createFont();
            fontTitle.setFontHeightInPoints((short) 18);
            fontTitle.setFontName("SansSerif");
            fontTitle.setBold(true);

            // CABECERA
            Font fontCellBold = workbook.createFont();
            fontCellBold.setFontHeightInPoints((short) 10);
            fontCellBold.setFontName("SansSerif");
            fontCellBold.setBold(true);

            // Merge Titulo
            createRowWithCells(sheet, 3, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 5, 22, 86));
            Row rowTitle = sheet.getRow(3);
            Cell cellTitle = rowTitle.getCell(22);
            cellTitle.setCellValue(titulo);

            // Merge Fecha
            createRowWithCells(sheet, 2, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 87, 100));
            Row rowFecha = sheet.getRow(2);
            Cell cellFecha = rowFecha.getCell(87);
            cellFecha.setCellValue("Fecha:");

            // Merge Hora
            createRowWithCells(sheet, 3, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 87, 100));
            Row rowHora = sheet.getRow(3);
            Cell cellHora = rowHora.getCell(87);
            cellHora.setCellValue("Hora:");

            // Merge FechaValue
            createRowWithCells(sheet, 2, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 101, 107));
            Row rowFechaValue = sheet.getRow(2);
            Cell cellFechaValue = rowFechaValue.getCell(101);
            cellFechaValue.setCellValue(sdfFecha.format(new Date()));

            // Merge HoraValue
            createRowWithCells(sheet, 3, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 101, 107));
            Row rowHoraValue = sheet.getRow(3);
            Cell cellHoraValue = rowHoraValue.getCell(101);
            cellHoraValue.setCellValue(sdfHora.format(new Date()));

            // Merge Numero Registros
            createRowWithCells(sheet, 7, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(7, 7, 87, 100));
            Row rowNumRecords = sheet.getRow(7);
            Cell cellNumRecords = rowNumRecords.getCell(87);
            cellNumRecords.setCellValue("Total de Registros:");

            // Merge Numero Registros Value
            createRowWithCells(sheet, 7, NUMERO_CELDAS);
            sheet.addMergedRegion(new CellRangeAddress(7, 7, 101, 107));
            Row rowValueNumRegistros = sheet.getRow(7);
            Cell cellValueNumRecords = rowValueNumRegistros.getCell(101);
            cellValueNumRecords.setCellValue(registros);

            // ESTILOS
            //TITULO
            CellStyle styleTitle = workbook.createCellStyle();
            styleTitle.setAlignment(HorizontalAlignment.CENTER);
            styleTitle.setVerticalAlignment(VerticalAlignment.CENTER);
            styleTitle.setFont(fontTitle);
            cellTitle.setCellStyle(styleTitle);

            // CABECERA
            XSSFCellStyle styleCabecera = (XSSFCellStyle) workbook.createCellStyle();
            styleCabecera.setBorderBottom(BorderStyle.THICK);
            styleCabecera.setBorderTop(BorderStyle.THICK);
            styleCabecera.setBorderLeft(BorderStyle.THICK);
            styleCabecera.setBorderRight(BorderStyle.THICK);
            styleCabecera.setFillForegroundColor(colorCabecera);
            styleCabecera.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            styleCabecera.setAlignment(HorizontalAlignment.CENTER);
            styleCabecera.setFont(fontCellBold);

            // FECHA
            XSSFCellStyle styleFecha = (XSSFCellStyle) workbook.createCellStyle();
            styleFecha.setAlignment(HorizontalAlignment.RIGHT);
            styleFecha.setFont(fontCellBold);
            cellFecha.setCellStyle(styleFecha);

            // FECHA_VALUE
            XSSFCellStyle styleFechaValue = (XSSFCellStyle) workbook.createCellStyle();
            styleFechaValue.setFont(fontCellBold);
            cellFechaValue.setCellStyle(styleFechaValue);

            // HORA
            XSSFCellStyle styleHora = (XSSFCellStyle) workbook.createCellStyle();
            styleHora.setAlignment(HorizontalAlignment.RIGHT);
            styleHora.setFont(fontCellBold);
            cellHora.setCellStyle(styleHora);

            // HORA_VALUE
            XSSFCellStyle styleHoraValue = (XSSFCellStyle) workbook.createCellStyle();
            styleHoraValue.setFont(fontCellBold);
            cellHoraValue.setCellStyle(styleHoraValue);

            // NUM_REGISTROS
            XSSFCellStyle styleNumRecords = (XSSFCellStyle) workbook.createCellStyle();
            styleNumRecords.setAlignment(HorizontalAlignment.RIGHT);
            styleNumRecords.setFont(fontCellBold);
            cellNumRecords.setCellStyle(styleNumRecords);

            // NUM_REGISTROS_VALUE
            XSSFCellStyle styleValueNumRecords = (XSSFCellStyle) workbook.createCellStyle();
            styleValueNumRecords.setAlignment(HorizontalAlignment.CENTER);
            styleValueNumRecords.setFont(fontCellBold);
            cellValueNumRecords.setCellStyle(styleValueNumRecords);

            // Merge Cabeceras
            createRowWithCells(sheet, 8, NUMERO_CELDAS, styleCabecera);
            mergeCellsInRowAndSet(sheet, 8, 3, NUMERO_CELDAS, listaCabecera.size(), listaCabecera, styleCabecera);

            //LOGO
            if (logoReporte != null) {

                int imgIdx = workbook.addPicture(logoReporte,Workbook.PICTURE_TYPE_PNG);

                CreationHelper helper = workbook.getCreationHelper();
                Drawing draw = sheet.createDrawingPatriarch();

                ClientAnchor clientAnchor = helper.createClientAnchor();
                clientAnchor.setCol1(1);
                clientAnchor.setRow1(1);

                Picture picture = draw.createPicture(clientAnchor, imgIdx);
                picture.resize(20, 5);
            }

            //Creacion de Registros
            int[] arrayIndex = getFirstColumns(NUMERO_CELDAS, numeroCabeceras);
            createRecords(sheet, arrayIndex, listaCabecera, listaAtributos, sdf);

            FileOutputStream stream = new FileOutputStream("C:\\data\\excel.xlsx");
            workbook.write(stream);
            workbook.close();
        } catch (Exception e) {
            System.out.println("Ocurrio un error al realizar el export a formato Excel. " + e);
        }
    }

    private static void createRowWithCells(Sheet sheet, int rowIndex, int cellCount) {
        int width = (int) (256 * (1.7));

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        for (int i = 0; i < cellCount; i++) {
            // Crear celda si no existe
            Cell cell = row.getCell(i + 1);
            if (cell == null) {
                cell = row.createCell(i + 1);
            }
            sheet.setColumnWidth(i + 1, width);
        }
    }

    private static void createRowWithCells(Sheet sheet, int rowIndex, int cellCount, CellStyle style) {
        int width = (int) (256 * (1.7));

        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        for (int i = 0; i < cellCount; i++) {
            // Crear celda si no existe
            Cell cell = row.getCell(i + 1);
            if (cell == null) {
                cell = row.createCell(i + 1);
                cell.setCellStyle(style);
            }

            sheet.setColumnWidth(i + 1, width);
        }
    }

    private static void mergeCellsInRowAndSet(Sheet sheet, int row, int firstCol, int NUMERO_CELDAS, int numeroCabeceras, List<String> values, CellStyle style) {
        int cellsGroup = (NUMERO_CELDAS - 2) / (numeroCabeceras - 1);
        int repartir = (NUMERO_CELDAS - 2) % (numeroCabeceras - 1);
        int lastCol = 0;

        sheet.addMergedRegion(new CellRangeAddress(row, row, 1, 2));
        Cell cell = sheet.getRow(row).getCell(1);
        cell.setCellValue(values.get(0));
        cell.setCellStyle(style);

        int valueIndex = 1;
        for (int i = 0; i < numeroCabeceras - 1; i++) {
            boolean hasExtraCell = (i != 0 && repartir >= 1) || (i == 0 && repartir > 0);

            lastCol = cellsGroup + firstCol + (hasExtraCell ? 0 : -1);
            sheet.addMergedRegion(new CellRangeAddress(row, row, firstCol, lastCol));

            if (valueIndex < values.size()) {
                cell = sheet.getRow(row).getCell(firstCol);
                cell.setCellValue(values.get(valueIndex));
                cell.setCellStyle(style);
                valueIndex++;
            }

            firstCol = lastCol + 1;

            if (hasExtraCell) {
                repartir--;
            }
        }
    }

    private static void createRecords(Sheet sheet, int[] arrayIndex, List<String> listaCabecera, List<List<Object>> listaAtributos, SimpleDateFormat sdf) {
        // Creación de Registros.
        int fila = 9;
        Workbook workbook = sheet.getWorkbook();
        CellStyle cellStyle = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName("SansSerif");
        font.setFontHeightInPoints((short) 8);
        cellStyle.setFont(font);
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        for (List<Object> listaTmp : listaAtributos) {
            mergeCellsInRow(sheet, fila, 1, 107, listaCabecera.size());
            Row rowReg = sheet.createRow(fila++);
            int i = 0;
            for (int j = 1; j <= arrayIndex.length; j++) {
                Cell cellReg = rowReg.createCell(arrayIndex[i]);
                cellReg.setCellStyle(cellStyle);
                Object obj = listaTmp.get(j - 1);
                if (obj instanceof Date) {
                    cellReg.setCellValue(sdf.format((Date) obj));
                } else if (obj instanceof Boolean) {
                    cellReg.setCellValue((Boolean) obj);
                } else if (obj instanceof String) {
                    cellReg.setCellValue((String) obj);
                } else if (obj instanceof Double) {
                    cellReg.setCellValue((Double) obj);
                } else if (obj instanceof Character) {
                    cellReg.setCellValue(((Character) obj).toString());
                } else if (obj instanceof Integer) {
                    cellReg.setCellValue(((Integer) obj).doubleValue());
                } else if (obj instanceof Long) {
                    cellReg.setCellValue(((Long) obj).doubleValue());
                }
                i++;
            }
        }
    }


    private static void mergeCellsInRow(Sheet sheet, int row, int firstCol, int NUMERO_CELDAS, int numeroCabeceras) {
        int cellsGroup = (NUMERO_CELDAS - 2) / (numeroCabeceras - 1);
        int repartir = (NUMERO_CELDAS - 2) % (numeroCabeceras - 1);
        int lastCol = 0;

        sheet.addMergedRegion(new CellRangeAddress(row, row, firstCol, firstCol + 1));
        firstCol += 2;

        for (int i = 0; i < numeroCabeceras - 1; i++) {
            boolean hasExtraCell = (i != 0 && repartir >= 1) || (i == 0 && repartir > 0);

            lastCol = cellsGroup + firstCol + (hasExtraCell ? 0 : -1);
            sheet.addMergedRegion(new CellRangeAddress(row, row, firstCol, lastCol));

            firstCol = lastCol + 1;

            if (hasExtraCell) {
                repartir--;
            }
        }
    }

    private static int[] getFirstColumns(int NUMERO_CELDAS, int numeroCabeceras) {
        int cellsGroup = (NUMERO_CELDAS - 2) / (numeroCabeceras - 1);
        int repartir = (NUMERO_CELDAS - 2) % (numeroCabeceras - 1);
        int lastCol = 0;
        int firstCol = 3;
        List<Integer> firstColumns = new ArrayList<>();
        firstColumns.add(1);

        for (int i = 0; i < numeroCabeceras - 1; i++) {
            boolean hasExtraCell = (i != 0 && repartir >= 1) || (i == 0 && repartir > 0);

            lastCol = cellsGroup + firstCol + (hasExtraCell ? 0 : -1);

            firstColumns.add(firstCol);

            firstCol = lastCol + 1;

            if (hasExtraCell) {
                repartir--;
            }
        }

        int[] firstColumnsArray = firstColumns.stream().mapToInt(Integer::intValue).toArray();
        return firstColumnsArray;
    }

    private static void fillColorInCells(Sheet sheet, List<List<Object>> listaRegistros) {
        // Crear estilo de celda con color de fondo blanco
        Workbook workbook = sheet.getWorkbook();
        CellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(IndexedColors.WHITE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // Aplicar el estilo a las celdas desde A1 hasta DC31
        for (int rowNum = 0; rowNum < listaRegistros.size() + 11; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                row = sheet.createRow(rowNum);
            }
            for (int colNum = 0; colNum < 1109; colNum++) {
                Cell cell = row.getCell(colNum);
                if (cell == null) {
                    cell = row.createCell(colNum);
                }
                cell.setCellStyle(style);
            }
        }
    }

}
