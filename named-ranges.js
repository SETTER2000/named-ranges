/**
 * Created by PhpStorm.
 * Company: Appalachian Ltd.
 * Developer: SETTER
 * Suite: appalachi.ru
 * Email: info@appalachi.ru
 * Date: 03.06.2017
 * Time: 11:03
 */
/**
 * Конструктор диапазона.
 * Именуем диапазоны входящего xlsx для удобного разбора и валидации
 * name - наименование диапазона
 * range - сам диапазон, например 'A1:K10'
 * style - общий стиль диапазона
 *
 */
function Ranges(workbook, name, range, style) {
    this.name = name;
    this.range = range;
    this.nameTwoColumn = 'C';
    this.pattern = /^\d+$/gi;
    this.style = style;
    this.error = {};
    this.currentError = 0;
    /**
     * Номера строк с ошибками. Причём если в строке 5 имеются 3 ошибки,
     * в разных столбцах, то строка 5 будет повторяться в массиве 3 раза.
     * @type {Array}
     */
    this.arrRowsError = [];
    this.arrRowsValid = [];
    this.workbook = workbook;
    /**
     * Всего колонок в книге
     */
    this.allColumns = workbook.sheet(0).usedRange()._numColumns;
    /**
     * Всего строк в книге
     */
    this.allRows = workbook.sheet(0).usedRange()._numRows;
    this.colorErrorCell = "ffc8ce";
}

/**
 * Функция возвращает уникальный массив на основе массива this.arrRowsError.
 * Массив будет содержать только уникальные номера строк с ошибками.
 * При этом массив this.arrRowsError сохраняет свои значения.
 *
 * @returns {Array}
 */
Ranges.prototype.uniqueArray = function () {
    var obj = {};
    // if(this.arrRowsError.length == allRows){
    //     res.badRequest({message:'Прайс не прошёл проверку! Ни одна из строк не может быть добавлена в общий прайс.'})
    // }
    for (var i = 0; i < this.arrRowsError.length; i++) {
        var str = this.arrRowsError[i];
        obj[str] = true; // запомнить строку в виде свойства объекта
    }
    return Object.keys(obj); // или собрать ключи перебором для IE8-
};


/**
 * Создать счётчик
 */
Ranges.prototype.createCounter = function () {
    function counter() {
        return counter.currentError++;
    }

    counter.currentError = this.currentError + 1;
    return counter;

};


/**
 * Получить диапазон объекта
 * @returns {*}
 */
Ranges.prototype.getRange = function () {
    return this.range;
};


/**
 * Получить имя диапазона
 * @returns {*}
 */
Ranges.prototype.getName = function () {
    return this.name;
};


/**
 * Валидация колонки по паттерну
 * @param pattern
 * @returns {number}
 */
Ranges.prototype.validationColumn = function (pattern) {
    let counter = this.createCounter();
    // Заменяем паттер, который был по умолчанию в классе
    if (pattern) this.pattern = pattern;

    // Проходим по всем ячейкам диапазона текужего объекта
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = `${range.value()}`;

        // Проверяем, если данные не прошли валидацию,
        // то красим ячейку красным цветом
        if (valueCell.match(this.pattern) != undefined) {
        } else {
            // кол-во ошибок
            this.currentError = counter();
            this.arrRowsError.push(range.rowNumber());
            //*********** !!! НЕ УДАЛЯТЬ! ***********************//
            //sails.log('rowNumber');
            //sails.log(range.rowNumber());
            //sails.log('row');
            //sails.log(range.row().cell(4).value());
            //sails.log('columnName');
            //sails.log(range.columnName() + '' + range.rowNumber());
            //sails.log('sheet');
            //sails.log(range.sheet().value());
            this.workbook.sheet(0).cell(currentCell).style("fill", this.colorErrorCell);
        }
    });

};


/**
 * Переводит строку в верхний регистр
 * @returns {string}
 */
Ranges.prototype.toUppCase = function () {
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = `${range.value()}`;
        if (valueCell !== 'undefined') {
            this.workbook.sheet(0).cell(currentCell).value(valueCell.toUpperCase());
        }


    });
};


/**
 * Режит строку длинее lengths
 * @returns {string}
 */
Ranges.prototype.toStringCut = function (lengths) {
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = `${range.value()}`;

        if (valueCell !== 'undefined') {
            this.workbook.sheet(0).cell(currentCell).value(valueCell.substring(0, lengths + 1));
        }
    });
};


/**
 * Валидация ячеек диапазона.
 * Ячейка должна содержать только одно значение из входящего массива
 * @param arr
 * @returns {number}
 */
Ranges.prototype.validationOneElementColumn = function (arr) {
    let counter = this.createCounter();
    // Проходим по всем ячейкам диапазона текужего объекта
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = `${range.value()}`;

        let e = valueCell.split(',');
        let e2 = valueCell.split(' ');

        // Проверяем, если данные не прошли валидацию,
        // то красим ячейку красным цветом
        if (e.length > 1 || e2.length > 1) {
            this.currentError = counter();
            this.arrRowsError.push(range.rowNumber());
            err = 1;
            //*********** !!! НЕ УДАЛЯТЬ! ***********************//
            //sails.log('rowNumber');
            //sails.log(range.rowNumber());
            //sails.log('row');
            //sails.log(range.row().cell(4).value());
            //sails.log('columnName');
            //sails.log(range.columnName() + '' + range.rowNumber());
            //sails.log('sheet');
            //sails.log(range.sheet().value());
            this.workbook.sheet(0).cell(currentCell).style("fill", this.this.colorErrorCell);
        } else {
            let val = e[0];
            if (val !== 'undefined') {
                if (arr.indexOf(val) > -1) {
                    this.workbook.sheet(0).cell(currentCell).value(val);
                } else {
                    this.currentError = counter();
                    this.arrRowsError.push(range.rowNumber());
                    this.workbook.sheet(0).cell(currentCell).style("fill", this.this.colorErrorCell);
                }
            }
        }
    });
};


/**
 * Валидация двух смежных ячеек на заполнение.
 * Впринципе можно использовать и для не смежных ячеек в одной строке.
 * @param nameTwoColumn
 * @returns {number}
 */
Ranges.prototype.validationUndefinedTwoColumn = function (nameTwoColumn) {
    let counter = this.createCounter();

    // Заменяем паттерн, который был по умолчанию в классе
    if (nameTwoColumn) this.nameTwoColumn = nameTwoColumn;

    // Проходим по всем ячейкам диапазона текужего объекта
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Координаты второй колонки
        let twoCell = this.nameTwoColumn + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = this.workbook.sheet(0).cell(currentCell).value();

        // Проверяем, если данные не прошли валидацию,
        // то красим ячейку красным цветом
        if (valueCell == undefined) {
            if (range.row().cell(this.nameTwoColumn).value() == undefined) {
                this.currentError = counter();
                this.arrRowsError.push(range.rowNumber());
                //*********** !!! НЕ УДАЛЯТЬ! ***********************//
                //sails.log('rowNumber');
                //sails.log(range.rowNumber());
                //sails.log('row');
                //sails.log(range.row().cell(4).value());
                //sails.log('columnName');
                //sails.log(range.columnName() + '' + range.rowNumber());
                //sails.log('sheet');
                //sails.log(range.sheet().value());
                this.workbook.sheet(0).cell(twoCell).style("fill", this.colorErrorCell);
                this.workbook.sheet(0).cell(currentCell).style("fill", this.colorErrorCell);
            }
        }
    });
};


/**
 * Валидация колонки на пустоту, т.е. ячейка не может быть пустой
 * @returns {number}
 */
Ranges.prototype.validationUndefinedColumn = function () {
    let counter = this.createCounter();
    // Проходим по всем ячейкам диапазона текужего объекта
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = this.workbook.sheet(0).cell(currentCell).value();

        // Проверяем, если данные не прошли валидацию,
        // то красим ячейку красным цветом
        if (valueCell == undefined) {
            this.currentError = counter();
            this.arrRowsError.push(range.rowNumber());
            //sails.log('COU:');
            //sails.log(this.currentError);
            this.workbook.sheet(0).cell(currentCell).style("fill", this.colorErrorCell);
        }
    });
};


/**
 * Удаляет из ячейки
 * @param pattern
 * @returns {number}
 */
Ranges.prototype.validationReplaceStringColumn = function (pattern, replace) {

    // Заменяем паттер, который был по умолчанию в классе
    if (pattern) this.pattern = pattern;

    // Проходим по всем ячейкам диапазона текужего объекта
    this.workbook.sheet(0).range(this.range).forEach(range => {

        // Координаты текущей ячейки. Например A3 или J55
        let currentCell = range.columnName() + '' + range.rowNumber();

        // Данные ячейки
        let valueCell = `${range.value()}`;

        if (valueCell.match(pattern)) {
            this.workbook.sheet(0).cell(currentCell).value(replace);
        }
        // Проверяем, если данные не прошли валидацию,
        // то красим ячейку красным цветом
        //if (valueCell.match(this.pattern) == undefined) {
        //    err=1;
        //    //*********** !!! НЕ УДАЛЯТЬ! ***********************//
        //    //sails.log('rowNumber');
        //    //sails.log(range.rowNumber());
        //    //sails.log('row');
        //    //sails.log(range.row().cell(4).value());
        //    //sails.log('columnName');
        //    //sails.log(range.columnName() + '' + range.rowNumber());
        //    //sails.log('sheet');
        //    //sails.log(range.sheet().value());
        //    this.workbook.sheet(0).cell(currentCell).style("fill", this.colorErrorCell);
        //}
    });
};


/**
 * Установить стили для диапазона
 */
Ranges.prototype.setStyle = function () {
    this.workbook.sheet(0).range(this.range).style(this.style);
};


/**
 * Собираем коллекцию объектов валидных строк из входящего прайса
 */
Ranges.prototype.setObjectRowsValid = function () {
    this.arrRowsValid = [];
    for (let i = 2; i <= allRows; i++) {
        if (all.arrRowsError.indexOf(i) < 0) {
            let o = {};
            for (let y = 1; y < 11; y++) {
                let nameColumnHeader = this.workbook.sheet(0).row(1).cell(y).value();
                //sails.log(nameColumnHeader);
                switch (nameColumnHeader) {
                    case 'ID':
                        nameColumnHeader = 'dax_id';
                        break;
                    case 'VendorID':
                        nameColumnHeader = 'vendor_id';
                        break;
                    case 'VendorID 2':
                        nameColumnHeader = 'vendor_id2';
                        break;
                    case 'Description':
                        nameColumnHeader = 'description';
                        break;
                    case 'Status':
                        nameColumnHeader = 'status';
                        break;
                    case 'Currency':
                        nameColumnHeader = 'currency';
                        break;
                    case 'DealerPrice':
                        nameColumnHeader = 'dealer_price';
                        break;
                    case 'SpecialPrice':
                        nameColumnHeader = 'special_price';
                        break;
                    case 'OpenPrice':
                        nameColumnHeader = 'open_price';
                        break;
                    case 'Note':
                        nameColumnHeader = 'note';
                        break;
                }
                o[nameColumnHeader] = this.workbook.sheet(0).row(i).cell(y).value();
            }
            if (o !== 'undefined') {
                o['vendor'] = vendor;
                //r.push(o);
                this.arrRowsValid.push(o);
            }
        }
    }
};


/**
 * Получить коллекцию объектов валидных строк входящего прайса
 * @returns {Array}
 */
Ranges.prototype.rowsValidArr = function () {
    this.setObjectRowsValid();
    return this.arrRowsValid;
};


/**
 * Кол-во корректных строк в книге
 */
Ranges.prototype.getAllValidCountRows = function () {
    return (this.allRows - all.uniqueArray().length);
};


/**
 * Какой процент ошибок в прайсе
 * @returns {number}
 */
Ranges.prototype.getAllErrorPercent = function () {
    let percent = this.arrRowsError.length * 100 / +((this.allRows - 1) * this.allColumns);
    return percent.toFixed(2);
};


/**
 * На сколько процентов прайс валидный. Проверка по ячейкам.
 * @returns {number}
 */
Ranges.prototype.getAllValidPercent = function () {
    let percent = 100 - this.getAllErrorPercent();
    return percent.toFixed(2);
};


module.exports = Ranges;