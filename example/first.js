/**
 * Created by PhpStorm.
 * Company: Appalachian Ltd.
 * Developer: SETTER
 * Suite: appalachi.ru
 * Email: info@appalachi.ru
 * Date: 03.06.2017
 * Time: 11:21
 */

const XlsxPopulate = require('xlsx-populate');
const Ranges = require('../named-ranges');
const DateRu = require('date-ru');


const pathToXlsxFile = 'Report 2017-05-31.xlsx';
var date = new Date();
var tpl = '%d.%m.%y %H:%M:%S';
var tpl2 = '%d.%m.%y_%H-%M-%S';
var dateRu = new DateRu(date, tpl);
var dateRu2 = new DateRu(date, tpl2);

/**
 * Текущая дата
 */
dateRu.localFormat();


/**
 * Именование статусов
 * @type {string}
 */
const statusSec = 'Принят частично';
const statusOk = 'Принят полностью';
const statusErr = 'Не принят';


/**
 * Формируем название файла отчёта NEW
 */
const date2 = dateRu.localFormat();
var fname = 'report-' + dateRu2 + '.xlsx';


/**
 * Путь и название файла отчета по загрузке
 * @type {string}
 */
const pathToReport = "./report/" + pathToXlsxFile;

/**
 *  Массив-шаблон названия столбцов,
 *  с его помощью будет проверяться соответствие столбцов в загружаемом файле
 * @type {string[]}
 */
const arrNameColumnsIdeal = [
    'Дата',
    'Отдел',
    'ФИО',
    'Таб. №',
    'События'
];

/**
 *
 * @param a
 * @returns {Array.<T>}
 */
Array.prototype.diff = function (a) {
    return this.filter(function (i) {
        return a.indexOf(i) < 0;
    });
};



// Загрузить существующую книгу xlsx
XlsxPopulate.fromFileAsync(pathToXlsxFile)
    .then(
        function (workbook) {
            "use strict";

            /**
             * Название листа.
             * @type {Sheet|Sheet|undefined}
             */
            const nameList = workbook.sheet(0);


            /**
             * Матрица вся книга.
             * @type {Range|undefined}
             */
            const matrix = workbook.sheet(0).usedRange();


            /**
             * Всего строк в книге
             */
            const allRows = matrix._numRows;


            /**
             * Проверяем наличие данных в книге
             */
            if (allRows < 2) {
                console.log('Книга пустая, проверять нечего! ');
                // return res.forbidden({
                //     message: 'Книга пустая, проверять нечего! '
                // });
            }

            
            
            /**
             * Количество колонок в прайсе, которое должно быть по умолчанию
             * @type {Number}
             */
            const countColumnsIdeal = arrNameColumnsIdeal.length;
            
            
            /**
             *  Массив для добавления имён столбцов из загружаемого файла
             * @type {Array}
             */
            const arrNameColumns = [];

            /**
             *  Получить названия колонок в загружаемом файле
             */
            for (var i = 1; i <= countColumnsIdeal; i++) {
                var nameColumn = workbook.sheet(0).row(7).cell(i).value();
                if (typeof nameColumn == 'undefined' && arrNameColumns.length < countColumnsIdeal) {
                    console.log('Кол-во колонок не совпадает с шаблоном по умолчанию! ');
                    // return res.forbidden({
                    //     message: 'Кол-во колонок не совпадает с шаблоном по умолчанию! ' +
                    //     'Должно быть ' + countColumnsIdeal + ' колонок. '
                    // });
                }
                if (arrNameColumnsIdeal[i - 1] !== nameColumn) {
                    console.log('Не верное имя колонки ++++!');
                    // return res.forbidden({
                    //     message: 'Не верное имя колонки ' +
                    //     nameColumn + '! Колонка должна называться ' + arrNameColumnsIdeal[i - 1]
                    // });
                }

                arrNameColumns.push(workbook.sheet(0).row(7).cell(i).value());
            }
            
            
            /**
             * Инициализация объектов диапазона
             * @type {Ranges}
             */
            // Объект all будет содержать общую информацию по книге
            const all = new Ranges(workbook,'ALL', `A1:J${allRows}`);
            const named = new Ranges(workbook,'NAMED', 'A1:F1');
            const info = new Ranges(workbook,'INFO', 'C3:D5');
            const header = new Ranges(workbook,'HEADER', 'A7:E7');
            const headerTwo = new Ranges(workbook,'HEADERTWO', 'E8:F8');
            const dateReport = new Ranges(workbook,'DATEREPORT', `A9:A${allRows}`);
            const department = new Ranges(workbook,'DEPARTMENT', `B9:B${allRows}`,{
                bold: true,
                fontFamily: 'Arial',
                numberFormat: 4,
                fontSize: 8,
                fontColor: 'ff0000',
                horizontalAlignment: 'center',
                verticalAlignment: 'center'
            });
            const fio = new Ranges(workbook,'FIO', `C9:C${allRows}`);
            const tab = new Ranges(workbook,'TAB', `D9:D${allRows}`);
            const coming = new Ranges(workbook,'COMING', `E9:E${allRows}`);
            const exit = new Ranges(workbook,'EXIT', `F9:F${allRows}`);


            /**
             * Инициализация имён диапазонов в загружаемой книге
             */

            // ALL
            workbook.definedName(all.getName(), workbook.sheet(0).range(all.getRange()));

            // NAMED
            workbook.definedName(named.getName(), workbook.sheet(0).range(named.getRange()));

            // INFO
            workbook.definedName(info.getName(), workbook.sheet(0).range(info.getRange()));

            // HEADER
            workbook.definedName(header.getName(), workbook.sheet(0).range(header.getRange()));

            // HEADERTWO
            workbook.definedName(headerTwo.getName(), workbook.sheet(0).range(headerTwo.getRange()));

            // DATEREPORT
            workbook.definedName(dateReport.getName(), workbook.sheet(0).range(dateReport.getRange()));

            // DEPARTMENT
            workbook.definedName(department.getName(), workbook.sheet(0).range(department.getRange()));

            // FIO
            workbook.definedName(fio.getName(), workbook.sheet(0).range(fio.getRange()));

            // TAB
            workbook.definedName(tab.getName(), workbook.sheet(0).range(tab.getRange()));

            // COMING
            workbook.definedName(coming.getName(), workbook.sheet(0).range(coming.getRange()));

            // EXIT
            workbook.definedName(exit.getName(), workbook.sheet(0).range(exit.getRange()));


            /**
             * VALIDATION
             */
            dateReport.validationColumn(/^[0-9]{4}-(0[1-9]|1[012])-(0[1-9]|1[0-9]|2[0-9]|3[01])|undefined/gi);
            fio.validationReplaceStringColumn(/([а-яё]+)\s(\(.*\))\s([а-яё]+)\s([а-яё]+)/gi, '$1 $3 $4');
            fio.validationColumn(/^([а-яё]+)\s([а-яё]+)\s([а-яё]+)|undefined/gi);
            coming.validationReplaceStringColumn(/((\d\d:\d\d) \(.*\))/gi, '$2');
            coming.validationReplaceStringColumn(/(\(нет\))/gi, 'undefined');
            exit.validationReplaceStringColumn(/((\d\d:\d\d) \(.*\))/gi, '$2');
            exit.validationReplaceStringColumn(/(\(нет\))/gi, 'undefined');


            /**
             * Применить стили для диапазона
             */
            department.setStyle();


            /**
             * Собираем номера строк, которые имеют ошибки во входящем прайсе
             * @type {Array.<T>}
             */
            all.arrRowsError = all.arrRowsError.concat(
                all.arrRowsError,
                header.arrRowsError,
                headerTwo.arrRowsError,
                named.arrRowsError,
                info.arrRowsError,
                dateReport.arrRowsError,
                department.arrRowsError,
                fio.arrRowsError,
                tab.arrRowsError,
                coming.arrRowsError,
                exit.arrRowsError
            );


            /**
             * Сверяет кол-во колонок с шаблоном по умолчанию
             */
            if (countColumnsIdeal != arrNameColumns.length) {
                console.log('Кол-во колонок не совпадает с шаблоном по умолчанию!');
                // return res.forbidden({
                //     message: 'Кол-во колонок не совпадает с шаблоном по умолчанию!'
                // });
            }

            /**
             * В загружаемом файле, проверяем соответствие заголовков столбцов шаблону и
             * возвращаем массив заголовков не соответствующих шаблону либо пустой массив
             *
             */
            var rs = arrNameColumnsIdeal.diff(arrNameColumns);
            console.log(rs);
            if (rs.length == 1) {
                const cll = workbook.sheet(0).row(7).find(rs[0]);
                workbook.sheet(0).row(7).cell(cll[0]._columnNumber).style({bold: true, fontColor: 'f90b0b'});
                workbook.toFileAsync(pathToReport);
                console.log('Ошибка в названии столбца ');
                // return res.badRequest({
                //     message: 'Ошибка в названии столбца ' + rs + '!',
                //     avatarFd: nameFileUpload,
                //     goReport: true
                // });
            }

            if (rs.length > 1) {
                console.log('Есть ошибки в названии столбцов ' + rs + '!');
                // return res.badRequest('Есть ошибки в названии столбцов ' + rs + '!');
            }


            // !!! НЕ УДАЛЯТЬ !!
            console.log('');
            console.log('***********************************');
            console.log('*    Всего ошибок                *');
            console.log('***********************************');
            console.log(all.arrRowsError.length);
            console.log('');
            console.log('***********************************');
            console.log('*    Кол-во ошибок в колонках     *');
            console.log('***********************************');
            console.log('Header: ' + header.currentError + ' Строки: ' + header.arrRowsError);
            console.log('HeaderTwo: ' + headerTwo.currentError + ' Строки: ' + headerTwo.arrRowsError);
            console.log('NAMED: ' + named.currentError + ' Строки: ' + named.arrRowsError);
            console.log('INFO: ' + info.currentError + ' Строки: ' + info.arrRowsError);
            console.log('DateReport: ' + dateReport.currentError + ' Строки: ' + dateReport.arrRowsError);
            console.log('Department: ' + department.currentError + ' Строки: ' + department.arrRowsError);
            console.log('Fio: ' + fio.currentError + ' Строки: ' + fio.arrRowsError);
            console.log('Tab: ' + tab.currentError + ' Строки: ' + tab.arrRowsError);
            console.log('Coming: ' + coming.currentError + ' Строки: ' + coming.arrRowsError);
            console.log('Exit: ' + exit.currentError + ' Строки: ' + exit.arrRowsError);
            console.log('');
            console.log('**********************************************');
            console.log('* Процентное соотношение загружаемого прайса *');
            console.log('**********************************************');
            console.log('Валидный на: ' + all.getAllValidPercent() + '%');
            console.log('Ошибок: ' + all.getAllErrorPercent() + '%');
            console.log('');


            if (all.uniqueArray().length == (allRows - 1)) {
                console.log('Книга не валидная! Ни одна строка не записана.');
                // return res.forbidden({
                //     message: 'Книга не валидная! Ни одна строка не записана.'
                // });
            }
           

            workbook.toFileAsync(pathToReport);

        }
    ).catch((error) => {
    console.log(error,'Promise error 88888');
});