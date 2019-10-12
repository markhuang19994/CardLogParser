package com

import com.fasterxml.jackson.databind.ObjectMapper
import org.apache.poi.ss.usermodel.CellStyle
import org.apache.poi.ss.usermodel.CreationHelper
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
        def logFiles = getRceLogFromDir(args[0] as File)
        def log = mergeRceLogFromDir(logFiles)
        def om = new ObjectMapper()

        def results = log.split('\n').toList().stream().filter {
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
            it.size() > 0 && it['ino'] && it['ino'] && it['CARD_TYPE'] && it['GIFTCODE']
        }.collect(Collectors.toList())

        println "找出${results.size()}筆資料:"

        //依照日期分類資料
        def dataCategoryByDateMap = [:]
        for (result in results) {
            def date = result['logTimeTrace'].toString().split(' ')[0]
            def temp = dataCategoryByDateMap[date]
            dataCategoryByDateMap[date] = temp ? temp << result : [result]
            println "時間：${result['logTimeTrace']}, 身份證：${result['ino']}, 卡號：${result['CARD_TYPE']}, REF_NO：${result['REF_NO']}, 禮物代號：${result['GIFTCODE']}"
        }

        for (entry in dataCategoryByDateMap.entrySet()) {
            createExcel(entry.key.toString(), entry.value as List)
        }

    }

    static List<File> getRceLogFromDir(File logDir) {
        def logName = 'CLM_WebLog_RCE2.log'
        logDir.listFiles({
            it.getName().matches("${logName}.*")
        } as FileFilter)?.sort { a, b ->
            def aName = a.getName()
            def bName = b.getName()
            if (aName == logName) return 1

            def aIdx = aName.replace(logName + '.', '') as int
            def bIdx = bName.replace(logName + '.', '') as int
            Integer.compare(aIdx, bIdx)
        }
    }

    static String mergeRceLogFromDir(List<File> logFiles) {
        StringBuilder sb = new StringBuilder()
        for (logFile in logFiles) {
            sb.append(logFile.text).append('\n')
        }
        sb.toString()
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
        for (datum in data) {
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

        //建立輸出流
        FileOutputStream fos = new FileOutputStream(new File(System.getProperty('user.dir'), excelName + '.xlsx'))
        workbook.write(fos)
        workbook.close()
        fos.close()
    }
}
