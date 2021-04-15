import javafx.scene.control.ComboBox
import javafx.scene.control.DatePicker
import javafx.scene.layout.BorderPane
import tornadofx.*
import java.time.LocalDate

// Автоматическое заполнения полей Код-Наименование в таблице (одно задано - другое заполнилось)
// Инкрементый поиск в выпадающих списках и полях таблицы.
// Выгрузка заполненного документа в Excel / Calc.

fun main(args: Array<String>) {
    launch<BebraApp>(args)
}

class BebraApp : App(BebraView::class, Styles::class)

class BebraView : View("Унифицированная форма N ОП-10") {
    override val root : BorderPane by fxml()

    private val actDatePicker : DatePicker by fxid()

    init {
        actDatePicker.value = LocalDate.now()
    }

    fun openSignatures() {
        SignaturesView().openWindow()
    }
}

class SignaturesView : View("Расшифровка подписей") {
    override val root : BorderPane by fxml()

    private val positionComboBox : ComboBox<String> by fxid()

    fun exit() = close()
}

class Styles : Stylesheet()