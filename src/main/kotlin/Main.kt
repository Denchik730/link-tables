import org.apache.poi.ss.usermodel.WorkbookFactory
import org.apache.commons.csv.CSVFormat
import org.apache.commons.csv.CSVParser
import org.apache.commons.csv.CSVPrinter
import java.io.File
import java.io.FileReader
import java.io.FileWriter
import java.util.regex.Pattern

// Функция для чтения CSV файла и создания мапы
fun readCsvToMap(filePath: String): Map<String, String> {
    val csvParser = CSVParser(FileReader(filePath), CSVFormat.DEFAULT.withDelimiter(';').withHeader())
    val result = mutableMapOf<String, String>()

    for (record in csvParser) {
        if (record.size() >= 2) {
            val testCaseId = record.get(0).trim()
            val teamName = record.get(1).trim()
            result[testCaseId] = teamName
        } else {
            println("Warning: Skipping invalid record - $record")
        }
    }
    csvParser.close()
    return result
}

// Функция для извлечения номеров тест-кейсов из строки
fun extractTestCaseIds(testCaseStr: String): List<String> {
    val regex = Pattern.compile("'([^']*)'")
    val matcher = regex.matcher(testCaseStr)
    val result = mutableListOf<String>()

    while (matcher.find()) {
        result.add(matcher.group(1).trim())
    }

    return result
}

fun main() {
    val excelFilePath = "./../../allure_report.xlsx" // Путь к вашему Excel файлу
    val csvFilePath = "./../../report.csv" // Путь к вашему CSV файлу
    val outputCsvFilePath = "./../../output.csv" // Путь к итоговому CSV файлу

    // Читаем данные из CSV файла и создаем мапу {тест-кейс: команда}
    val csvData = readCsvToMap(csvFilePath)
    println("CSV Data: $csvData")

    // Открываем Excel файл
    val workbook = WorkbookFactory.create(File(excelFilePath))
    val sheet = workbook.getSheetAt(0)

    // Открываем CSV файл для записи
    val writer = FileWriter(outputCsvFilePath)
    val csvPrinter = CSVPrinter(writer, CSVFormat.DEFAULT.withHeader("Test Case IDs", "GitLab Link", "Team Name"))

    // Итерируемся по строкам Excel файла и добавляем данные из CSV
    val testCaseToTeamMap = mutableMapOf<String, String>()
    val testCaseToGitLabLinkMap = mutableMapOf<String, String>()

    for (rowIndex in 1 until sheet.physicalNumberOfRows) {
        val row = sheet.getRow(rowIndex)
        val testCaseCell = row?.getCell(0)
        val gitLabLinkCell = row?.getCell(1)

        val testCaseStr = testCaseCell?.stringCellValue?.trim()
        val gitLabLink = gitLabLinkCell?.stringCellValue?.trim()

        println("Processing row $rowIndex: testCaseStr=$testCaseStr, gitLabLink=$gitLabLink")

        if (testCaseStr != null) {
            // Извлекаем номера тест-кейсов из строки
            val testCaseIds = extractTestCaseIds(testCaseStr)

            // Для каждого номера тест-кейса устанавливаем ссылку на GitLab
            for (testCaseId in testCaseIds) {
                testCaseToGitLabLinkMap[testCaseId] = gitLabLink ?: "Link not available"

                // Устанавливаем команду для каждого номера тест-кейса
                val teamName = csvData[testCaseId] ?: "No team found"
                testCaseToTeamMap[testCaseStr] = teamName
            }
        } else {
            println("Warning: Empty test case string at row $rowIndex")
        }
    }

    // Записываем результаты в итоговый CSV файл
    for ((testCaseIds, teamName) in testCaseToTeamMap) {
        val gitLabLink = testCaseToGitLabLinkMap.values.firstOrNull() ?: "Link not available"
        csvPrinter.printRecord(testCaseIds, gitLabLink, teamName)
        println("Written to CSV: $testCaseIds, $gitLabLink, $teamName")
    }

    // Закрываем ресурсы
    csvPrinter.flush()
    csvPrinter.close()
    writer.close()
    workbook.close()

    println("Данные успешно объединены и сохранены в $outputCsvFilePath")
}
