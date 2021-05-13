import com.ibm.icu.text.RuleBasedNumberFormat
import javafx.scene.control.*
import javafx.scene.layout.BorderPane
import javafx.util.Callback
import javafx.util.converter.BigDecimalStringConverter
import javafx.util.converter.IntegerStringConverter
import org.apache.poi.ss.usermodel.WorkbookFactory
import tornadofx.*
import java.io.File
import java.io.FileOutputStream
import java.math.BigDecimal
import java.math.RoundingMode
import java.time.LocalDate
import java.time.format.DateTimeFormatter
import java.util.*

// Автоматическое заполнения полей Код-Наименование в таблице (одно задано - другое заполнилось)
// Инкрементый поиск в выпадающих списках и полях таблицы.
// Выгрузка заполненного документа в Excel / Calc.

fun main(args: Array<String>) {
    launch<BebraApp>(args)
}

class BebraApp : App(BebraView::class, Styles::class)

class BebraView : View("Унифицированная форма N ОП-10") {
    override val root: BorderPane by fxml()

    private val workbook = WorkbookFactory.create(File("op10.xls"))

    private val actNumberField: TextField by fxid()
    private val actDatePicker: DatePicker by fxid()

    private val organization: ComboBox<String> by fxid()
    private val orgs = listOf(
        "ООО Рога и Копыта",
        "МРК Вектор",
    ).asObservable()

    private val subdivision: ComboBox<String> by fxid()
    private val subs = listOf(
        "Столовая №1",
        "Столовая №2",
        "Столовая №3",
    ).asObservable()

    private val okpo1: TextField by fxid()
    private val okpo2: TextField by fxid()
    private val okpd: TextField by fxid()
    private val operationKind: TextField by fxid()

    private val specPer: TextField by fxid()
    private val specRub: TextField by fxid()
    private val specKop: TextField by fxid()

    private val saltPer: TextField by fxid()
    private val saltRub: TextField by fxid()
    private val saltKop: TextField by fxid()

    private val totalRub: TextField by fxid()
    private val totalKop: TextField by fxid()

    private val totalRefRub: TextField by fxid()
    private val totalRefKop: TextField by fxid()

    init {
        specPer.text = "0"
        saltPer.text = "0"

        specPer.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = specPer.text.toFloatOrNull()
                if (v == null || v < 0) {
                    specPer.text = "0"
                }
            }
        }

        saltPer.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = saltPer.text.toFloatOrNull()
                if (v == null || v < 0) {
                    saltPer.text = "0"
                }
            }
        }
    }

    init {
        saltRub.text = "0"
        saltKop.text = "0"
        specRub.text = "0"
        specKop.text = "0"
        totalRub.text = "0"
        totalKop.text = "0"

        totalRefRub.isDisable = true
        totalRefRub.text = "0"

        totalRefKop.isDisable = true
        totalRefKop.text = "0"

        specRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = specRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    specRub.text = "0"
                } else {
                    totalRefRub.text = (v + (saltRub.text.toIntOrNull() ?: 0) + ((saltKop.text.toIntOrNull()
                        ?: 0) + (specKop.text.toIntOrNull() ?: 0)) / 100).toString()
                }
            }
        }

        specKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = specKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    specKop.text = "0"
                } else {
                    specKop.text = (v % 100).toString()
                    totalRefKop.text = ((v + (saltKop.text.toIntOrNull() ?: 0)) % 100).toString()
                    totalRefRub.text = ((specRub.text.toIntOrNull() ?: 0) + (saltRub.text.toIntOrNull()
                        ?: 0) + (v + (saltKop.text.toIntOrNull() ?: 0)) / 100).toString()
                }
            }
        }

        saltRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = saltRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    saltRub.text = "0"
                } else {
                    totalRefRub.text = (v + (specRub.text.toIntOrNull() ?: 0) + ((saltKop.text.toIntOrNull()
                        ?: 0) + (specKop.text.toIntOrNull() ?: 0)) / 100).toString()
                }
            }
        }

        saltKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = saltKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    saltKop.text = "0"
                } else {
                    saltKop.text = (v % 100).toString()
                    totalRefKop.text = ((v + (specKop.text.toIntOrNull() ?: 0)) % 100).toString()
                    totalRefRub.text = ((specRub.text.toIntOrNull() ?: 0) + (saltRub.text.toIntOrNull()
                        ?: 0) + (v + (specKop.text.toIntOrNull() ?: 0)) / 100).toString()
                }
            }
        }
    }


    private val attRub: TextField by fxid()
    private val attKop: TextField by fxid()

    init {
        attRub.text = "0"
        attKop.text = "0"

        attRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = attRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    attRub.text = "0"
                }
            }
        }

        attKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = attKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    attKop.text = "0"
                } else {
                    attKop.text = (v % 100).toString()
                }
            }
        }

        totalRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = totalRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    totalRub.text = "0"
                }
            }
        }

        totalKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = totalKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    totalKop.text = "0"
                } else {
                    totalKop.text = (v % 100).toString()
                }
            }
        }
    }

    private val nalRub: TextField by fxid()
    private val nalKop: TextField by fxid()

    private val vyrRub: TextField by fxid()
    private val vyrKop: TextField by fxid()

    init {
        nalRub.text = "0"
        nalKop.text = "0"
        vyrRub.text = "0"
        vyrKop.text = "0"

        nalRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = nalRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    nalRub.text = "0"
                }
            }
        }

        nalKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = nalKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    nalKop.text = "0"
                } else {
                    nalKop.text = (v % 100).toString()
                }
            }
        }

        vyrRub.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = vyrRub.text.toIntOrNull()
                if (v == null || v < 0) {
                    vyrRub.text = "0"
                }
            }
        }

        vyrKop.focusedProperty().addListener { _, _, newValue ->
            if (!newValue) {
                val v = vyrKop.text.toIntOrNull()
                if (v == null || v < 0) {
                    vyrKop.text = "0"
                } else {
                    vyrKop.text = (v % 100).toString()
                }
            }
        }
    }

    private val naklad: TextField by fxid()
    private val zabor: TextField by fxid()

    private val table: TableView<TableRecord> by fxid()

    private val idColumn: TableColumn<TableRecord, Int?> by fxid()
    private val nameColumn: TableColumn<TableRecord, String?> by fxid()
    private val codeColumn: TableColumn<TableRecord, String?> by fxid()
    private val sellPriceColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val cashQuantityColumn: TableColumn<TableRecord, Int?> by fxid()
    private val cashSumColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val buffetQuantityColumn: TableColumn<TableRecord, Int?> by fxid()
    private val buffetSumColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val employeeQuantityColumn: TableColumn<TableRecord, Int?> by fxid()
    private val employeeSumColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val totalQuantityColumn: TableColumn<TableRecord, Int?> by fxid()
    private val totalSumColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val productionPriceColumn: TableColumn<TableRecord, BigDecimal?> by fxid()
    private val productionSumColumn: TableColumn<TableRecord, BigDecimal?> by fxid()

    private val product = object {
        private val entries = listOf(
            "Test 1" to "030",
            "Test 2" to "322",
            "Test 3" to "451",
            "Test 4" to "671",
        )

        val namesToSellPrice = listOf(
            "Test 1" to 10,
            "Test 2" to 20,
            "Test 3" to 30,
            "Test 4" to 40,
        ).associate { it.first to it.second.toBigDecimal() }


        val namesToProdPrice = listOf(
            "Test 1" to 5,
            "Test 2" to 10,
            "Test 3" to 15,
            "Test 4" to 20,
        ).associate { it.first to it.second.toBigDecimal() }

        val entriesMap = entries.toMap()
        val entriesMapInverted = entries.associate { (k, v) -> v to k }

        val names = entries.map { it.first }.asObservable()
        val codes = entries.map { it.second }.asObservable()
    }

    init {
        organization.value = orgs.first()
        organization.items = orgs

        subdivision.value = subs.first()
        subdivision.items = subs

        actDatePicker.value = LocalDate.now()

        idColumn.cellValueFactory = Callback { observable(it.value, TableRecord::id) }

        nameColumn.isEditable = true
        nameColumn.cellValueFactory = Callback { observable(it.value, TableRecord::name) }
        nameColumn.useComboBox(product.names)
        val commitNameColumn = nameColumn.onEditCommit
        nameColumn.setOnEditCommit {
            it.rowValue.code = product.entriesMap[it.newValue]!!

            it.rowValue.sellPrice = product.namesToSellPrice[it.newValue]!!
            it.rowValue.productionPrice = product.namesToProdPrice[it.newValue]!!

            recalculateData(it.rowValue)

            it.tableView.refresh()
            commitNameColumn.handle(it)
        }

        codeColumn.isEditable = true
        codeColumn.cellValueFactory = Callback { observable(it.value, TableRecord::code) }
        codeColumn.useComboBox(product.codes)
        val commitCodeColumn = codeColumn.onEditCommit
        codeColumn.setOnEditCommit {
            it.rowValue.name = product.entriesMapInverted[it.newValue]!!

            it.rowValue.sellPrice = product.namesToSellPrice[it.rowValue.name]!!
            it.rowValue.productionPrice = product.namesToProdPrice[it.rowValue.name]!!

            recalculateData(it.rowValue)

            it.tableView.refresh()
            commitCodeColumn.handle(it)
        }

        sellPriceColumn.isEditable = true
        sellPriceColumn.cellValueFactory = Callback { observable(it.value, TableRecord::sellPrice) }
        sellPriceColumn.useTextField(BigDecimalStringConverter())
        val commitSellPrice = sellPriceColumn.onEditCommit
        sellPriceColumn.setOnEditCommit {
            recalculateData(it.rowValue)
            it.tableView.refresh()
            commitSellPrice.handle(it)
        }

        cashQuantityColumn.isEditable = true
        cashQuantityColumn.cellValueFactory = Callback { observable(it.value, TableRecord::cashQuantity) }
        cashQuantityColumn.useTextField(IntegerStringConverter())
        val commitCashQuantity = cashQuantityColumn.onEditCommit
        cashQuantityColumn.setOnEditCommit {
            commitCashQuantity.handle(it)
            recalculateData(it.rowValue)
            it.tableView.refresh()
        }

        cashSumColumn.cellValueFactory = Callback { observable(it.value, TableRecord::cashSum) }

        buffetQuantityColumn.isEditable = true
        buffetQuantityColumn.cellValueFactory = Callback { observable(it.value, TableRecord::buffetQuantity) }
        buffetQuantityColumn.useTextField(IntegerStringConverter())
        val commitBuffetQuantity = buffetQuantityColumn.onEditCommit
        buffetQuantityColumn.setOnEditCommit {
            commitBuffetQuantity.handle(it)
            recalculateData(it.rowValue)
            it.tableView.refresh()
        }

        buffetSumColumn.cellValueFactory = Callback { observable(it.value, TableRecord::buffetSum) }

        employeeQuantityColumn.isEditable = true
        employeeQuantityColumn.cellValueFactory = Callback { observable(it.value, TableRecord::employeeQuantity) }
        employeeQuantityColumn.useTextField(IntegerStringConverter())
        val commitEmployeeQuantity = employeeQuantityColumn.onEditCommit
        employeeQuantityColumn.setOnEditCommit {
            commitEmployeeQuantity.handle(it)
            recalculateData(it.rowValue)
            it.tableView.refresh()
        }

        employeeSumColumn.cellValueFactory = Callback { observable(it.value, TableRecord::employeeSum) }

        totalQuantityColumn.cellValueFactory = Callback { observable(it.value, TableRecord::totalQuantity) }
        totalSumColumn.cellValueFactory = Callback { observable(it.value, TableRecord::totalSum) }

        productionPriceColumn.isEditable = true
        productionPriceColumn.cellValueFactory = Callback { observable(it.value, TableRecord::productionPrice) }
        productionPriceColumn.useTextField(BigDecimalStringConverter())
        val commitProductionPrice = productionPriceColumn.onEditCommit
        productionPriceColumn.setOnEditCommit {
            commitProductionPrice.handle(it)
            recalculateData(it.rowValue)
            it.tableView.refresh()
        }

        productionSumColumn.cellValueFactory = Callback { observable(it.value, TableRecord::productionSum) }

        table.contextMenu = contextmenu {
            item("Добавить строку").action {
                addItemToTable()
            }
            val deleteButton = item("Удалить строку") {
                action {
                    table.selectedItem?.let {
                        table.items.remove(it)
                        table.items.forEachIndexed { index, tableRecord -> tableRecord.id = index + 1 }
                    }
                }
            }
            setOnShown { deleteButton.isDisable = table.selectedItem == null }
        }
        addItemToTable()
    }

    private fun addItemToTable() {
        table.items.add(
            TableRecord(
                table.items.size + 1,
                product.names.first(),
                product.codes.first(),
                product.namesToSellPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                product.namesToProdPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN)
            )
        )
    }

    private fun recalculateData(record: TableRecord) {
        record.cashSum = record.cashQuantity.toBigDecimal() * record.sellPrice
        record.buffetSum = record.buffetQuantity.toBigDecimal() * record.sellPrice
        record.employeeSum = record.employeeQuantity.toBigDecimal() * record.sellPrice
        record.totalQuantity = record.cashQuantity + record.buffetQuantity + record.employeeQuantity

        val totalQuantity = record.totalQuantity.toBigDecimal()
        record.totalSum = totalQuantity * record.sellPrice
        record.productionSum = totalQuantity * record.productionPrice
    }

    var rukPod: String = ""
    var rukTransc: String = ""
    var zavPod: String = ""
    var zavTransc: String = ""
    var marPod: String = ""
    var marTransc: String = ""
    var buhPod: String = ""
    var buhTransc: String = ""
    var casPod: String = ""
    var casTransc: String = ""
    var rukPos = "Директор"

    fun openSignatures() = SignaturesView(this).openWindow()

    private fun numberToMonth(n: Int): String {
        return when (n) {
            1 -> "января"
            2 -> "февраля"
            3 -> "марта"
            4 -> "апреля"
            5 -> "мая"
            6 -> "июня"
            7 -> "июля"
            8 -> "августа"
            9 -> "сентября"
            10 -> "октября"
            11 -> "ноября"
            12 -> "декабря"
            else -> ""
        }
    }

    fun exportToExcel() {
        val sheet = workbook.getSheetAt(0)
        // Заполнение кодов
        sheet.getRow(5).getCell(68).setCellValue(okpo1.text)
        sheet.getRow(6).getCell(68).setCellValue(okpo2.text)
        sheet.getRow(8).getCell(68).setCellValue(okpd.text)
        sheet.getRow(9).getCell(68).setCellValue(operationKind.text)

        // Заполнение номера документа и даты
        sheet.getRow(13).getCell(42).setCellValue(actNumberField.text)
        sheet.getRow(13).getCell(50).setCellValue(actDatePicker.value.format(DateTimeFormatter.ofPattern("dd.MM.yyyy")))
        sheet.getRow(16).getCell(62).setCellValue(actDatePicker.value.format(DateTimeFormatter.ofPattern("dd")))
        sheet.getRow(16).getCell(64).setCellValue(numberToMonth(actDatePicker.value.monthValue))
        sheet.getRow(16).getCell(72).setCellValue(actDatePicker.value.format(DateTimeFormatter.ofPattern("yyyy")))

        // Заполнение организации
        sheet.getRow(5).getCell(0).setCellValue(organization.value)
        // Заполнение структурного подразделения
        sheet.getRow(7).getCell(0).setCellValue(subdivision.value)

        // Заполнение подписей
        sheet.getRow(12).getCell(61).setCellValue(rukPos)
        sheet.getRow(14).getCell(60).setCellValue(rukPod)
        sheet.getRow(14).getCell(67).setCellValue(rukTransc)
        sheet.getRow(68).getCell(14).setCellValue(zavPod)
        sheet.getRow(68).getCell(23).setCellValue(zavTransc)
        sheet.getRow(70).getCell(7).setCellValue(marPod)
        sheet.getRow(70).getCell(20).setCellValue(marTransc)
        sheet.getRow(72).getCell(0).setCellValue(rukPos)
        sheet.getRow(72).getCell(11).setCellValue(rukPod)
        sheet.getRow(72).getCell(20).setCellValue(rukTransc)
        sheet.getRow(83).getCell(12).setCellValue(buhPod)
        sheet.getRow(83).getCell(21).setCellValue(buhTransc)

        // Заполнение приложения
        sheet.getRow(79).getCell(8).setCellValue(naklad.text)
        sheet.getRow(81).getCell(11).setCellValue(naklad.text)

        // Заполнение справки и прочего
        sheet.getRow(61).getCell(21).setCellValue(specPer.text)
        sheet.getRow(61).getCell(36).setCellValue(specRub.text)
        sheet.getRow(61).getCell(54).setCellValue(specKop.text)

        sheet.getRow(63).getCell(20).setCellValue(saltPer.text)
        sheet.getRow(63).getCell(36).setCellValue(saltRub.text)
        sheet.getRow(63).getCell(54).setCellValue(saltKop.text)

        sheet.getRow(65).getCell(36).setCellValue(totalRefRub.text)
        sheet.getRow(65).getCell(54).setCellValue(totalRefKop.text)

        val f = RuleBasedNumberFormat(Locale.forLanguageTag("ru"), RuleBasedNumberFormat.SPELLOUT)

        sheet.getRow(57).getCell(0).setCellValue(f.format(attRub.text.toInt()))
        sheet.getRow(57).getCell(67).setCellValue(f.format(attKop.text.toInt()))

        sheet.getRow(58).getCell(30).setCellValue(f.format(totalRub.text.toInt()))
        sheet.getRow(58).getCell(67).setCellValue(f.format(totalKop.text.toInt()))

        sheet.getRow(75).getCell(8).setCellValue(f.format(vyrRub.text.toInt()))
        sheet.getRow(75).getCell(69).setCellValue(f.format(vyrKop.text.toInt()))

        sheet.getRow(79).getCell(59).setCellValue(nalRub.text)
        sheet.getRow(79).getCell(69).setCellValue(nalKop.text)

        // Заполнение строк
        for ((i, v) in table.items.withIndex()) {
            val row = sheet.getRow(if (i <= 10) (26 + i) else (35 + i))

            row.getCell(0).setCellValue(v.id.toString())
            row.getCell(4).setCellValue(v.name)
            row.getCell(15).setCellValue(v.code)
            row.getCell(18).setCellValue(v.sellPrice.toString())
            row.getCell(23).setCellValue(v.cashQuantity.toString())
            row.getCell(27).setCellValue(v.cashSum.toString())
            row.getCell(32).setCellValue(v.buffetQuantity.toString())
            row.getCell(36).setCellValue(v.buffetSum.toString())
            row.getCell(41).setCellValue(v.employeeQuantity.toString())
            row.getCell(45).setCellValue(v.employeeSum.toString())
            row.getCell(58).setCellValue(v.totalQuantity.toString())
            row.getCell(62).setCellValue(v.totalSum.toString())
            row.getCell(66).setCellValue(v.productionPrice.toString())
            row.getCell(71).setCellValue(v.productionSum.toString())
        }

        fun sumRecords(acc: TableRecord, tableRecord: TableRecord) = acc.apply {
            cashQuantity += tableRecord.cashQuantity
            cashSum += tableRecord.cashSum
            buffetQuantity += tableRecord.buffetQuantity
            buffetSum += tableRecord.buffetSum
            employeeQuantity += tableRecord.employeeQuantity
            employeeSum += tableRecord.employeeSum
            totalQuantity += tableRecord.totalQuantity
            totalSum += tableRecord.totalSum
            productionSum += tableRecord.productionSum
        }

        val firstElevenItems = table.items.take(11).fold(
            TableRecord(
                table.items.size + 1,
                product.names.first(),
                product.codes.first(),
                product.namesToSellPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                product.namesToProdPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN)
            )
        ) { acc, tableRecord ->
            sumRecords(acc, tableRecord)
        }

        val lastItems = table.items.drop(11).fold(
            TableRecord(
                table.items.size + 1,
                product.names.first(),
                product.codes.first(),
                product.namesToSellPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                product.namesToProdPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN)
            )
        ) { acc, tableRecord ->
            sumRecords(acc, tableRecord)
        }

        val allItems = table.items.fold(
            TableRecord(
                table.items.size + 1,
                product.names.first(),
                product.codes.first(),
                product.namesToSellPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                0,
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN),
                product.namesToProdPrice[product.names.first()]!!.setScale(2, RoundingMode.HALF_DOWN),
                BigDecimal(0).setScale(2, RoundingMode.HALF_DOWN)
            )
        ) { acc, tableRecord ->
            sumRecords(acc, tableRecord)
        }

        sheet.getRow(37).getCell(23).setCellValue(firstElevenItems.cashQuantity.toString())
        sheet.getRow(37).getCell(27).setCellValue(firstElevenItems.cashSum.toString())
        sheet.getRow(37).getCell(32).setCellValue(firstElevenItems.buffetQuantity.toString())
        sheet.getRow(37).getCell(36).setCellValue(firstElevenItems.buffetSum.toString())
        sheet.getRow(37).getCell(41).setCellValue(firstElevenItems.employeeQuantity.toString())
        sheet.getRow(37).getCell(45).setCellValue(firstElevenItems.employeeSum.toString())
        sheet.getRow(37).getCell(58).setCellValue(firstElevenItems.totalQuantity.toString())
        sheet.getRow(37).getCell(62).setCellValue(firstElevenItems.totalSum.toString())
        sheet.getRow(37).getCell(71).setCellValue(firstElevenItems.productionSum.toString())

        sheet.getRow(53).getCell(23).setCellValue(lastItems.cashQuantity.toString())
        sheet.getRow(53).getCell(27).setCellValue(lastItems.cashSum.toString())
        sheet.getRow(53).getCell(32).setCellValue(lastItems.buffetQuantity.toString())
        sheet.getRow(53).getCell(36).setCellValue(lastItems.buffetSum.toString())
        sheet.getRow(53).getCell(41).setCellValue(lastItems.employeeQuantity.toString())
        sheet.getRow(53).getCell(45).setCellValue(lastItems.employeeSum.toString())
        sheet.getRow(53).getCell(58).setCellValue(lastItems.totalQuantity.toString())
        sheet.getRow(53).getCell(62).setCellValue(lastItems.totalSum.toString())
        sheet.getRow(53).getCell(71).setCellValue(lastItems.productionSum.toString())

        sheet.getRow(54).getCell(23).setCellValue(allItems.cashQuantity.toString())
        sheet.getRow(54).getCell(27).setCellValue(allItems.cashSum.toString())
        sheet.getRow(54).getCell(32).setCellValue(allItems.buffetQuantity.toString())
        sheet.getRow(54).getCell(36).setCellValue(allItems.buffetSum.toString())
        sheet.getRow(54).getCell(41).setCellValue(allItems.employeeQuantity.toString())
        sheet.getRow(54).getCell(45).setCellValue(allItems.employeeSum.toString())
        sheet.getRow(54).getCell(58).setCellValue(allItems.totalQuantity.toString())
        sheet.getRow(54).getCell(62).setCellValue(allItems.totalSum.toString())
        sheet.getRow(54).getCell(71).setCellValue(allItems.productionSum.toString())

        with(FileOutputStream("bebrus.xls")) {
            workbook.write(this)
        }
    }
}

class TableRecord(
    var id: Int,
    var name: String,
    var code: String,
    var sellPrice: BigDecimal,
    var cashQuantity: Int,
    var cashSum: BigDecimal,
    var buffetQuantity: Int,
    var buffetSum: BigDecimal,
    var employeeQuantity: Int,
    var employeeSum: BigDecimal,
    var totalQuantity: Int,
    var totalSum: BigDecimal,
    var productionPrice: BigDecimal,
    var productionSum: BigDecimal,
)

class SignaturesView(private val bebra: BebraView) : View("Расшифровка подписей") {
    override val root: BorderPane by fxml()

    private val positionComboBox: ComboBox<String> by fxid()
    private val positions = listOf("Директор", "Заместитель директора по общественному питанию").asObservable()

    init {
        positionComboBox.value = bebra.rukPos
        positionComboBox.items = positions
    }

    private val rukPod: TextField by fxid()
    private val rukTransc: TextField by fxid()

    private val zavPod: TextField by fxid()
    private val zavTransc: TextField by fxid()

    private val marPod: TextField by fxid()
    private val marTransc: TextField by fxid()

    private val buhPod: TextField by fxid()
    private val buhTransc: TextField by fxid()

    private val casPod: TextField by fxid()
    private val casTransc: TextField by fxid()

    init {
        rukPod.text = bebra.rukPod
        rukTransc.text = bebra.rukTransc

        zavPod.text = bebra.zavPod
        zavTransc.text = bebra.zavTransc

        marPod.text = bebra.marPod
        marTransc.text = bebra.marTransc

        buhPod.text = bebra.buhPod
        buhTransc.text = bebra.buhTransc

        casPod.text = bebra.casPod
        casTransc.text = bebra.casTransc
    }

    fun exit() = close()
    fun saveAndExit() {
        bebra.rukPod = rukPod.text
        bebra.rukTransc = rukTransc.text

        bebra.zavPod = zavPod.text
        bebra.zavTransc = zavTransc.text

        bebra.marPod = marPod.text
        bebra.marTransc = marTransc.text

        bebra.buhPod = buhPod.text
        bebra.buhTransc = buhTransc.text

        bebra.casPod = casPod.text
        bebra.casTransc = casTransc.text

        bebra.rukPos = positionComboBox.value
        println(bebra.rukPos)

        close()
    }
}

class Styles : Stylesheet()