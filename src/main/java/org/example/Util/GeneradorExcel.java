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
import java.util.*;

public class GeneradorExcel {
    final static int NUMERO_COLUMNAS = 107;

    public static void exportExcel(String pestana, String titulo, int registros, List<String> listaCabecera, List<List<Object>> listaAtributos) throws FileNotFoundException {
        SimpleDateFormat sdf = new SimpleDateFormat("HH:mm:ss dd/MM/yyyy");
        SimpleDateFormat sdfHora = new SimpleDateFormat("HH:mm:ss");
        SimpleDateFormat sdfFecha = new SimpleDateFormat("dd/MM/yyyy");
        final int NUMERO_COLUMNAS = 107;
        int numeroCabeceras = listaCabecera.size();

        try {
            SXSSFWorkbook workbook = new SXSSFWorkbook();
            Sheet sheet = workbook.createSheet(pestana);

            // en el codigo del BO no será necesario
            InputStream is = new FileInputStream("C:\\data\\logo-ripley.png");
            byte[] logoReporte = IOUtils.toByteArray(is);
            byte[] celeste = new byte[]{(byte) 0, (byte) 255, (byte) 255};
            XSSFColor colorCabecera = new XSSFColor(celeste);
            int[] arrayIndex = getFirstColumns(NUMERO_COLUMNAS, numeroCabeceras);
            setWidth(sheet, NUMERO_COLUMNAS);

            // FONDO BLANCO
            //fillColorInCells(sheet, listaAtributos);

            //FUENTE
            // TITULO
            Font fontTitle = workbook.createFont();
            fontTitle.setFontHeightInPoints((short) 18);
            fontTitle.setFontName("SansSerif");
            fontTitle.setBold(true);

            // CABECERA SUBTITULOS
            Font fontCellBold = workbook.createFont();
            fontCellBold.setFontHeightInPoints((short) 10);
            fontCellBold.setFontName("SansSerif");
            fontCellBold.setBold(true);

            // Merge Titulo
            createRowWithCells(sheet, 3, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 5, 22, 86));
            Row rowTitle = sheet.getRow(3);
            Cell cellTitle = rowTitle.getCell(22);
            cellTitle.setCellValue(titulo);

            // Merge Fecha
            createRowWithCells(sheet, 2, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 87, 100));
            Row rowFecha = sheet.getRow(2);
            Cell cellFecha = rowFecha.getCell(87);
            cellFecha.setCellValue("Fecha:");

            // Merge Hora
            createRowWithCells(sheet, 3, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 87, 100));
            Row rowHora = sheet.getRow(3);
            Cell cellHora = rowHora.getCell(87);
            cellHora.setCellValue("Hora:");

            // Merge FechaValue
            createRowWithCells(sheet, 2, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(2, 2, 101, 107));
            Row rowFechaValue = sheet.getRow(2);
            Cell cellFechaValue = rowFechaValue.getCell(101);
            cellFechaValue.setCellValue(sdfFecha.format(new Date()));

            // Merge HoraValue
            createRowWithCells(sheet, 3, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(3, 3, 101, 107));
            Row rowHoraValue = sheet.getRow(3);
            Cell cellHoraValue = rowHoraValue.getCell(101);
            cellHoraValue.setCellValue(sdfHora.format(new Date()));

            // Merge Numero Registros
            createRowWithCells(sheet, 7, NUMERO_COLUMNAS);
            sheet.addMergedRegion(new CellRangeAddress(7, 7, 87, 100));
            Row rowNumRecords = sheet.getRow(7);
            Cell cellNumRecords = rowNumRecords.getCell(87);
            cellNumRecords.setCellValue("Total de Registros:");

            // Merge Numero Registros Value
            createRowWithCells(sheet, 7, NUMERO_COLUMNAS);
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
            styleCabecera.setBorderBottom(BorderStyle.MEDIUM);
            styleCabecera.setBorderTop(BorderStyle.MEDIUM);
            styleCabecera.setBorderLeft(BorderStyle.MEDIUM);
            styleCabecera.setBorderRight(BorderStyle.MEDIUM);
            styleCabecera.setFillForegroundColor(colorCabecera);
            styleCabecera.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            styleCabecera.setAlignment(HorizontalAlignment.CENTER);
            styleCabecera.setVerticalAlignment(VerticalAlignment.CENTER);
            styleCabecera.setFont(fontCellBold);
            styleCabecera.setWrapText(true);

            // REGISTROS
            CellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFont(fontCellBold);
            cellStyle.setBorderBottom(BorderStyle.THIN);
            cellStyle.setBorderTop(BorderStyle.THIN);
            cellStyle.setBorderLeft(BorderStyle.THIN);
            cellStyle.setBorderRight(BorderStyle.THIN);
            cellStyle.setAlignment(HorizontalAlignment.CENTER);
            cellStyle.setWrapText(true);

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
            createRowWithCells(sheet, 8, styleCabecera, arrayIndex, listaCabecera);
            mergeCellsInRow(sheet, 8, arrayIndex, NUMERO_COLUMNAS);
            //con este código se realiza el ajuste automático
            int column =0;
            int height = calculateLineBreaks(sheet,8,column, arrayIndex);
            Row row = sheet.getRow(8);
            row.setHeightInPoints(height * sheet.getDefaultRowHeightInPoints());

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
            createRecords(sheet, arrayIndex, listaCabecera, listaAtributos, sdf, cellStyle);

            FileOutputStream stream = new FileOutputStream("C:\\data\\excel.xlsx");
            workbook.write(stream);
            workbook.close();
        } catch (Exception e) {
            System.out.println("Ocurrio un error al realizar el export a formato Excel. " + e);
        }
    }

    private static void createRowWithCells(Sheet sheet, int rowIndex, int cellCount) {
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
        }
    }

    private static void setWidth(Sheet sheet, int numberColumn) {
        int width = (int) (256 * (1.7));
        for (int i = 0; i < numberColumn; i++) {
            sheet.setColumnWidth(i + 1, width);
        }
    }

    //TODO:CORREGIR MÉTODO
    private static void createRowWithCells(Sheet sheet, int rowIndex, CellStyle style, int[] column, List<String> cabeceras) {
        Row row = sheet.createRow(rowIndex);
            //row.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());
        for (int i = 0; i < column.length; i++) {
            Cell cell = row.createCell(column[i]);
            cell.setCellValue(cabeceras.get(i));
            cell.setCellStyle(style);
        }
    }

    // Calcular saltos de linea de una fila específica
    private static int calculateLineBreaks(Sheet sheet, int row, int column, int[] arrayIndex) {
        int charsWithCell;
        int charsLengthCell;

        charsWithCell = getNumCharsInMergedCell(sheet, row, column);
        charsLengthCell = getKeyOfMaxValue(mapOfMergedCellLengths(arrayIndex, sheet, row));
        double division = (double) charsLengthCell / charsWithCell;

        return (int) Math.round(division);
    }

    //Obtener llave(indice de la columna) del valor máximo
    private static Integer getKeyOfMaxValue(Map<Integer, Integer> map) {
        Optional<Map.Entry<Integer, Integer>> maxEntry = map.entrySet()
                .stream()
                .max(Map.Entry.comparingByValue());
        return maxEntry.map(Map.Entry::getKey).orElse(null);
    }

    // Mapa de los valores de celdas combinadas
    private static Map<Integer, Integer> mapOfMergedCellLengths(int[] arrayIndex, Sheet sheet, int row) {
        Map<Integer, Integer> cellWeightsInRow = new HashMap<>();
        for (int index : arrayIndex) {
            int weight = getLengthOfMergedCell(sheet, row, index);
            cellWeightsInRow.put(index, weight);
        }
        return cellWeightsInRow;
    }

    // Obtener el valor de una celda combinada
    private static int getLengthOfMergedCell(Sheet sheet, int row, int column) {
        DataFormatter formatter = new DataFormatter();
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(row, column)) {
                Row firstRow = sheet.getRow(range.getFirstRow());
                Cell firstCellOfMergedRegion = firstRow.getCell(range.getFirstColumn());
                String cellValue = formatter.formatCellValue(firstCellOfMergedRegion);
                return cellValue.length();
            }
        }
        return -1;
    }

    private static int getNumCharsInMergedCell(Sheet sheet, int row, int column) {
        DataFormatter formatter = new DataFormatter();
        for (CellRangeAddress range : sheet.getMergedRegions()) {
            if (range.isInRange(row, column)) {
                Row firstRow = sheet.getRow(range.getFirstRow());
                Cell firstCellOfMergedRegion = firstRow.getCell(range.getFirstColumn());
                String cellValue = formatter.formatCellValue(firstCellOfMergedRegion);
                return cellValue.length();
            }
        }
        return -1;
    }

    private static void mergeCellsInRow(Sheet sheet, int row, int[] arrayIndex, int numCeldas) {
        for(int i=0; i< arrayIndex.length-1; i++) {
            if (i == arrayIndex.length - 2) {
                sheet.addMergedRegion(new CellRangeAddress(row, row, arrayIndex[i], arrayIndex[i+1]-1));
                sheet.addMergedRegion(new CellRangeAddress(row, row, arrayIndex[i+1], numCeldas));
            } else {
                sheet.addMergedRegion(new CellRangeAddress(row, row, arrayIndex[i], arrayIndex[i+1]-1));

            }
        }
    }

    private static void createRecords(Sheet sheet, int[] arrayIndex, List<String> listaCabecera, List<List<Object>> listaAtributos, SimpleDateFormat sdf, CellStyle cellStyle) {
        // Creación de Registros.
        int fila = 9;
        for (List<Object> listaTmp : listaAtributos) {
            Row rowReg = sheet.createRow(fila++);
            //rowReg.setHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());
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
        fila = 9;
        for(int i=0; i<listaCabecera.size(); i++) {
            mergeCellsInRowByCol(sheet, fila, arrayIndex, listaAtributos.size(), NUMERO_COLUMNAS, i);
        }
    }

    private static void mergeCellsInRowByCol(Sheet sheet, int row, int[] arrayIndex, int numRegistros, int numCeldas, int columna) {
        for(int i = row; i< row + numRegistros; i++) {
            if(columna == arrayIndex.length - 2) {
                sheet.addMergedRegion(new CellRangeAddress(i, i, arrayIndex[columna], arrayIndex[columna+1]-1));
                sheet.addMergedRegion(new CellRangeAddress(i, i, arrayIndex[columna + 1], numCeldas));
            } else if(columna < arrayIndex.length - 1) {
                sheet.addMergedRegion(new CellRangeAddress(i, i, arrayIndex[columna], arrayIndex[columna + 1] - 1));
            }
        }
    }

    private static int[] getFirstColumns(int NUMERO_COLUMNAS, int numeroCabeceras) {
        int cellsGroup = (NUMERO_COLUMNAS - 2) / (numeroCabeceras - 1);
        int repartir = (NUMERO_COLUMNAS - 2) % (numeroCabeceras - 1);
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
