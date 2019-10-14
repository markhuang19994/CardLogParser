package com

import org.apache.poi.ss.usermodel.Cell
import org.apache.poi.ss.usermodel.CellType
import org.apache.poi.ss.usermodel.DateUtil
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.sql.Timestamp
import java.text.SimpleDateFormat

class ExcelCitiDataReader implements CitiDataReader {

    List<CitiData> read(File dataFile) {
        ByteArrayInputStream bis = new ByteArrayInputStream(dataFile.bytes)
        Workbook workbook = new XSSFWorkbook(bis)
        Sheet sheet = workbook.getSheetAt(0)

        List<CitiData> resultList = new ArrayList<>()
        for (int i = 1; i <= sheet.getLastRowNum(); i++) {
            def row = sheet.getRow(i)
            if (row == null || row.getLastCellNum() == (short) -1) continue

            //crd_typ
            String crdType = getStringCellValue(row.getCell(1))

            //dte_entry
            String date = new SimpleDateFormat('yyyy-MM-dd').format(row.getCell(3).getDateCellValue())

            //crd_typ
            String idno = getStringCellValue(row.getCell(6))

            resultList << new CitiData(
                    crdType: crdType.trim(),
                    date: date.trim(),
                    id: idno.trim(),
                    rowNum: i
            )
        }

        resultList
    }

    String getStringCellValue(Cell cell) {
        return String.valueOf(getCellValue(cell))
    }

    private Object getCellValue(Cell cell) {
        Object cellValue = null
        switch (cell.getCellType()) {
            case CellType.STRING:
                cellValue = cell.getStringCellValue()
                break

            case CellType.FORMULA:
                cell.setCellType(CellType.STRING)
                cellValue = cell.getStringCellValue()
                break

            case CellType.NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    cellValue = cell.getDateCellValue().toString()
                } else {
                    cellValue = Double.toString(cell.getNumericCellValue())
                }
                break

            case CellType.BLANK:
                cellValue = ""
                break

            case CellType.BOOLEAN:
                cellValue = Boolean.toString(cell.getBooleanCellValue())
                break

        }
        return cellValue
    }

    Timestamp getCellAsTimestamp(Cell cell) {
        Date dateCellValue = cell.getDateCellValue()
        if (dateCellValue == null) return null
        return new Timestamp(dateCellValue.getTime())
    }

    Double parseDouble(String cellValue) {
        if (cellValue == null || "".equals(cellValue)) {
            return null
        } else {
            return Double.valueOf(cellValue)
        }
    }
}