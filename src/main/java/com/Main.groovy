package com

import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
import org.apache.poi.ss.usermodel.Sheet
import org.apache.poi.ss.usermodel.Workbook
import org.apache.poi.xssf.usermodel.XSSFRow
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import javax.script.ScriptEngine
import javax.script.ScriptEngineManager
import java.text.SimpleDateFormat
import java.util.stream.Collectors

/**
 * @author MarkHuang* @version
 * <ul>
 *  <li>10/9/19, MarkHuang,new
 * </ul>
 * @since 10/9/19
 */
class Main {
    static ScriptEngineManager scriptEngineManager = new ScriptEngineManager()
    static ScriptEngine nashorn = scriptEngineManager.getEngineByName("nashorn")

    static void main(String[] args) {
        def citiExcelFile = new File(System.getProperty('user.dir'), 'input.xlsx')
        def citiData = new ExcelCitiDataReader().read(citiExcelFile)

        def logDirPath = new File(System.getProperty('user.dir'), 'log')
        def logFiles = getRceLogFromDirAll(logDirPath)
        def log = mergeRceLogFromDir(logFiles)

        def data = getRequestDataInLog(log)
        data.forEach({
            it['ino'] = it['ino'] ?: it['idNO'] ?: it['ID_NO']
        })

        def distData = distinctRefNo(data)
        printFindData(distData)

        def categoryDataMap = categoryDataByDate(distData)

        def citiNeedDataMap = [:]
        def citiExcelMap = [:]

        citiData.parallelStream().forEach { citiDatum ->
            def date = citiDatum.date
            def id = citiDatum.id
            def dataList = categoryDataMap[date]
            if (dataList) {
                def dataFilterById = dataList.findAll {
                    (it['ino'] as List)[0] == id && it['CARD_TYPE'] && (it['CARD_TYPE'] as List)[0].toString().contains(citiDatum.crdType)
                }
                if (dataFilterById.size() > 0) {
                    citiExcelMap.put(citiDatum.rowNum, dataFilterById.collect {
                        new GiftCodeAndTimeTrace(
                                giftCode: it['GIFTCODE'][0],
                                timeTrace: it['logTimeTrace']
                        )
                    })
                    citiNeedDataMap[date] = citiNeedDataMap[date] ?: []
                    (citiNeedDataMap[date] as List).addAll(dataFilterById)
                } else {
                    citiExcelMap.put(citiDatum.rowNum, null)
                    println "無法過濾的資料:${date} ${id},id不存在..."
                }
            } else {
                citiExcelMap.put(citiDatum.rowNum, null)
                println "無法過濾的資料:${date} ${id},日期不存在..."
            }
        }

        writeDataToExcel(citiNeedDataMap)
        writeDataToExcel2(citiExcelFile, citiExcelMap)
        println '執行完畢，產出結果在result資料夾。'
    }

    static getRequestDataInLog(String log) {
        def om = new ObjectMapper()

        log.split('\n').toList().parallelStream().filter {
            it.matches('^2019-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2},\\d+(.*)?Request Data:[\\s\\S]*$')
        }.map {
            def timeAndData = it.replaceAll('^(2019-\\d{2}-\\d{2} \\d{2}:\\d{2}:\\d{2},\\d+).*?Request Data: ([\\s\\S]*)', '$1@@$2')
            timeAndData.split('@@')
        }.map {
            if (it.length < 2 || it[1] == null || it[1] == '') return [:]
            try {
                def map = om.readValue(it[1], Map.class)
                map['logTimeTrace'] = it[0]
                return map
            } catch (Exception e) {
                try {
                    String json = nashorn.eval("""
                        var a = ${it[1]}     
                        a['fail'] = ''    
                        JSON.stringify(a)
                    """.toString())
                    def map = om.readValue(json, Map.class)
                    map['logTimeTrace'] = it[0]
                    return map
                } catch (Exception e2) {
                    e.printStackTrace()
                    e2.printStackTrace()
                }
            }
            [:]
        }.filter {
            it.size() > 0 && (it['ino'] || it['idNO'] || it['ID_NO']) && it['CARD_TYPE'] && it['GIFTCODE']
        }.collect(Collectors.toList())
    }

    static List<File> getRceLogFromDirAll(File logDirParent) {
        List<File> logDirs = logDirParent.listFiles({
            it.getName().matches("^2019-\\d{2}-\\d{2}\$")
        } as FileFilter)?.sort { a, b ->
            def aName = a.getName()
            def bName = b.getName()

            def x = aName.replaceAll('^2019-(\\d{2}-\\d{2})', '$1').split('-')
            def y = bName.replaceAll('^2019-(\\d{2}-\\d{2})', '$1').split('-')

            def scX = (x[0] as int) * 100 + (x[1] as int)
            def scY = (y[0] as int) * 100 + (y[1] as int)
            Integer.compare(scX, scY)
        }?.toList() ?: []

        def result = []
        for (logDir in logDirs) {
            result.addAll(getRceLogFromDir(logDir))
        }
        result
    }

    static List<File> getRceLogFromDir(File logDir) {
        def logName = 'CLM_WebLog_RCE2.log'
        logDir.listFiles({
            it.getName().matches("${logName}.*")
        } as FileFilter)?.sort { a, b ->
            def aName = a.getName()
            def bName = b.getName()
            if (aName == logName) return -1
            if (bName == logName) return 1

            def aIdx = aName.replace(logName + '.', '') as int
            def bIdx = bName.replace(logName + '.', '') as int
            Integer.compare(aIdx, bIdx)
        }
    }

    /**
     * 如果有多個重複的refsNo，則以時間最晚的那一個為準
     * @param data
     * @return
     */
    static distinctRefNo(List data) {
        def newData = []
        def refNoList = []
        for (datum in data.reverse()) {
            def refNo = datum['REF_NO']
            if (refNo && !(refNo in refNoList)) {
                newData << datum
                refNoList << refNo
            }
        }
        newData.reverse()
    }

    static String mergeRceLogFromDir(List<File> logFiles) {
        StringBuilder sb = new StringBuilder()
        for (logFile in logFiles) {
            sb.append(logFile.text).append('\n')
        }
        sb.toString()
    }

    static categoryDataByDate(data) {
        def dataCategoryByDateMap = [:]
        for (datum in data) {
            def date = datum['logTimeTrace'].toString().split(' ')[0]
            def temp = dataCategoryByDateMap[date]
            dataCategoryByDateMap[date] = temp ? temp << datum : [datum]
        }
        dataCategoryByDateMap
    }

    static printFindData(data) {
        println "找出${data.size()}筆資料:"
        for (datum in data) {
            println "時間：${datum['logTimeTrace']}, 身份證：${datum['ino']}, 卡號：${datum['CARD_TYPE']}, REF_NO：${datum['REF_NO']}, 禮物代號：${datum['GIFTCODE']}"
        }
    }

    static writeDataToExcel(dataMap) {
        for (entry in dataMap.entrySet()) {
            createExcel(entry.key.toString(), entry.value as List)
        }
    }

    static createExcel(String excelName, List data) {
        XSSFWorkbook workbook = new XSSFWorkbook()
        XSSFSheet sheet = workbook.createSheet("report")
        XSSFRow headRow = sheet.createRow(0)
        headRow.createCell(0).setCellValue('日期')
        headRow.createCell(1).setCellValue('身分證字號')
        headRow.createCell(2).setCellValue('卡號')
        headRow.createCell(3).setCellValue('REF_NO')
        headRow.createCell(4).setCellValue('禮物代號')

        def sdf = new SimpleDateFormat('yyyy-MM-dd HH:mm:ss,S')
        CellStyle cellStyle = workbook.createCellStyle()
        CreationHelper createHelper = workbook.getCreationHelper()
        cellStyle.setDataFormat(
                createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm:ss,S"))

        int i = 1
        for (datum in (data as List<Map<String, List>>)) {
            XSSFRow row = sheet.createRow(i)
            def cell0 = row.createCell(0,)
            cell0.setCellValue(sdf.parse(datum['logTimeTrace']?.toString()))
            cell0.setCellStyle(cellStyle)
            row.createCell(1).setCellValue(datum['ino'][0]?.toString())
            row.createCell(2).setCellValue(datum['CARD_TYPE'][0]?.toString())
            row.createCell(3).setCellValue(datum['REF_NO'][0]?.toString())
            row.createCell(4).setCellValue(datum['GIFTCODE'][0]?.toString())
            i++
        }

        sheet.autoSizeColumn(0)
        sheet.autoSizeColumn(1)
        sheet.autoSizeColumn(2)
        sheet.autoSizeColumn(3)
        sheet.autoSizeColumn(4)

        FileOutputStream fos = new FileOutputStream(new File(System.getProperty('user.dir'), 'result/' + excelName + '.xlsx'))
        workbook.write(fos)
        workbook.close()
        fos.close()
    }

    static writeDataToExcel2(File dataFile, Map citiExcelMap) {
        ByteArrayInputStream bis = new ByteArrayInputStream(dataFile.bytes)
        Workbook workbook = new XSSFWorkbook(bis)
        Sheet sheet = workbook.getSheetAt(0)

        sheet.getRow(0).createCell(11).setCellValue('detail')
        citiExcelMap.forEach { rowNum, giftCodeAndTimeTraceList ->
            def row = sheet.getRow(rowNum as int)

            def cell2 = row.getCell(2) ?: row.createCell(2)
            if (giftCodeAndTimeTraceList == null) {
                cell2.setCellValue('未找到')
            } else {
                def cell2Sb = new StringBuilder()
                def cell11Sb = new StringBuilder()
                for (giftCodeAndTimeTrace in (giftCodeAndTimeTraceList as List<GiftCodeAndTimeTrace>)) {
                    cell2Sb.append(giftCodeAndTimeTrace.giftCode).append(', ')
                    cell11Sb.append(giftCodeAndTimeTrace.timeTrace + '=>' + giftCodeAndTimeTrace.giftCode).append('| ')
                }

                def cell2Val = cell2Sb.toString()
                def cell11Val = cell11Sb.toString()
                cell2.setCellValue(cell2Val[0..cell2Val.length() - 3])
                row.createCell(11).setCellValue(cell11Val[0..cell11Val.length() - 3])
            }
        }

        FileOutputStream fos = new FileOutputStream(new File(System.getProperty('user.dir'), 'result/result.xlsx'))
        workbook.write(fos)
        workbook.close()
        fos.close()
    }

    static class GiftCodeAndTimeTrace {
        String giftCode
        String timeTrace
    }
}
