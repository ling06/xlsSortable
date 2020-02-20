var XlsSortable = function (selector, options) {
    options = Object.assign({
        autostart: true,
    }, options || {});

    let guessFunctions = [
        function (value, column) {
            if (/^\d+[.\-\/]\d+[.\-\/]\d+ \d+:\d+(:\d+)?$/.test(value)) {
                return 'datetime';
            }
            return false;
        },
        function (value, column) {
            if (/^\d+[.\-\/]\d+[.\-\/]\d+$/.test(value)) {
                return 'date';
            }
            return false;
        },
        function (value, column) {
            if (/^\d+:\d+(:\d+)?$/.test(value)) {
                return 'time';
            }
            return false;
        },
        function (value, column) {
            if (!isNaN(parseFloat(value))) {
                return 'number';
            }
            return false;
        },
        function (value, column) {
            return 'string';
        },
    ];

    let sortFunctions = {
        datetime: function (a, b, sortDir) {
            a = a.replace(/(\d+)[.\/\-](\d+)[.\/\-](\d+)/, '$3-$2-$1');
            b = b.replace(/(\d+)[.\/\-](\d+)[.\/\-](\d+)/, '$3-$2-$1');
            return (sortDir === 'asc' && a > b) || (sortDir === 'desc' && a < b) ?
                1 :
                (
                    (sortDir === 'asc' && a < b) || (sortDir === 'desc' && a > b) ?
                        -1 :
                        0
                );
        },
        date: function (a, b, sortDir) {
            a = a.replace(/(\d+)[.\/\-](\d+)[.\/\-](\d+)/, '$3-$2-$1');
            b = b.replace(/(\d+)[.\/\-](\d+)[.\/\-](\d+)/, '$3-$2-$1');
            return (sortDir === 'asc' && a > b) || (sortDir === 'desc' && a < b) ?
                1 :
                (
                    (sortDir === 'asc' && a < b) || (sortDir === 'desc' && a > b) ?
                        -1 :
                        0
                );
        },
        time: function (a, b, sortDir) {
            return (sortDir === 'asc' && a > b) || (sortDir === 'desc' && a < b) ?
                1 :
                (
                    (sortDir === 'asc' && a < b) || (sortDir === 'desc' && a > b) ?
                        -1 :
                        0
                );
        },
        number: function (a, b, sortDir) {
            a = parseFloat(a);
            b = parseFloat(b);
            return (sortDir === 'asc' && a > b) || (sortDir === 'desc' && a < b) ?
                1 :
                (
                    (sortDir === 'asc' && a < b) || (sortDir === 'desc' && a > b) ?
                        -1 :
                        0
                );
        },
        string: function (a, b, sortDir) {
            return (sortDir === 'asc' && a > b) || (sortDir === 'desc' && a < b) ?
                1 :
                (
                    (sortDir === 'asc' && a < b) || (sortDir === 'desc' && a > b) ?
                        -1 :
                        0
                );
        },
    };

    let printFunctions = {
        datetime: function (value) {
            if (/(\d+)\-(\d+)\-(\d+)/.test(value)) {
                value = value.replace(/(\d+)\-(\d+)\-(\d+)/, '$3.$2.$1');
            }
            if (/\d+\/\d+\/\d+/.test(value)) {
                value = value.split('/').join('.');
            }
            return value;
        },
        date: function (value) {
            if (/\d+\-\d+\-\d+/.test(value)) {
                return value.split('-').reverse().join('.');
            }
            if (/\d+\/\d+\/\d+/.test(value)) {
                return value.split('/').join('.');
            }
            return value;
        },
        time: function (value) {
            return value;
        },
        number: function (value) {
            return value;
        },
        string: function (value) {
            return value;
        },
    };

    let queue = [];

    let parseExcel = function(file, callback) {
        let reader = new FileReader();
        reader.onload = (event) => {
            let data = event.target.result;
            let workbook = XLSX.read(data, {type: 'binary'});
            workbook.SheetNames.forEach(function(sheetName) {
                var XLRowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]);
                if (typeof callback === 'function') {
                    callback(XLRowObject);
                }
            });
        };
        reader.onerror = function(ex) {
            console.log(ex);
        };
        reader.readAsBinaryString(file);
    };

    function loadScript(url, callback) {
        var head = document.head,
            script = document.createElement('script');
        script.type = 'text/javascript';
        script.src = url;
        script.onreadystatechange = callback;
        script.onload = callback;
        head.appendChild(script);
    }

    function loadTable(url, callback) {
        fetch(url, {cache: 'no-cache'})
            .then(response => response.blob())
            .then(blob => parseExcel(blob, callback));
    }

    function numToRow(num, str) {
        str = str || '';
        const codeA = 'A'.charCodeAt(0);
        const code = num % 26;
        str = String.fromCharCode(codeA + code) + str;
        num -= code;
        return num ? numToRow(num/26-1, str) : str;
    }

    function guessType(value, column) {
        for (let i = 0, il = guessFunctions.length; i < il; ++i) {
            let type = guessFunctions[i](value, column);
            if (type) {
                return type;
            }
        }
        return 'string';
    }

    function sortData(data, sortBy, sortDir) {
        data.body = data.body.sort((a, b) => {
            if (typeof data.minValues[sortBy] !== 'undefined') {
                if (a[sortBy] === data.minValues[sortBy] && b[sortBy] !== data.minValues[sortBy]) {
                    return sortDir === 'asc' ? -1 : 1;
                }
                if (a[sortBy] !== data.minValues[sortBy] && b[sortBy] === data.minValues[sortBy]) {
                    return sortDir === 'asc' ? 1 : -1;
                }
                if (a[sortBy] === data.minValues[sortBy] && b[sortBy] === data.minValues[sortBy]) {
                    return 0;
                }
            }
            if (typeof data.maxValues[sortBy] !== 'undefined') {
                if (a[sortBy] === data.maxValues[sortBy] && b[sortBy] !== data.maxValues[sortBy]) {
                    return sortDir === 'asc' ? 1 : -1;
                }
                if (a[sortBy] !== data.maxValues[sortBy] && b[sortBy] === data.maxValues[sortBy]) {
                    return sortDir === 'asc' ? -1 : 1;
                }
                if (a[sortBy] === data.maxValues[sortBy] && b[sortBy] === data.maxValues[sortBy]) {
                    return 0;
                }
            }
            return sortFunctions[data.types[sortBy]](a[sortBy], b[sortBy], sortDir);
        });
        return data.body;
    }

    function formTable(tableElement, tableData, options) {
        options = Object.assign({
            sortBy: -1,
            sortDir: '',
        }, options || {});
        if (options.sortBy !== -1 && options.sortDir) {
            tableData.body = sortData(tableData, options.sortBy, options.sortDir);
        }

        let table = document.createElement('table');
        table.classList.add(options.tableClass);
        if (options.theme) {
            table.classList.add(options.tableClass + '_theme_' + options.theme);
        }

        let thead = document.createElement('thead');
        let theadTr = document.createElement('tr');
        for (let i = 0, il = tableData.headers.length; i < il; ++i) {
            let th = document.createElement('th');
            let colName = numToRow(i);
            let isColSortable = typeof options.noSort === 'undefined' || options.noSort.indexOf(colName) === -1;
            if (isColSortable) {
                th.classList.add(options.tableSortClass);
                if (options.sortBy === i) {
                    th.classList.add(options.tableSortClass + '-' + options.sortDir);
                }
                (i => {
                    th.addEventListener('click', () => {
                        if (options.sortBy !== i) {
                            options.sortDir = 'asc';
                        } else {
                            options.sortDir = options.sortDir === 'asc' ?
                                'desc' :
                                'asc';
                        }
                        options.sortBy = i;
                        sortData(tableData, options.sortBy, options.sortDir);
                        formTable(tableElement, tableData, options);
                        fireEvent('tableSorted', options);
                    });
                })(i);
            }
            th.innerText = tableData.headers[i];
            theadTr.appendChild(th);
        }
        thead.appendChild(theadTr);
        table.appendChild(thead);

        let tbody = document.createElement('tbody');
        tableData.body.forEach(row => {
            let tbodyTr = document.createElement('tr');
            row.forEach(cell => {
                let td = document.createElement('td');
                td.innerText = cell;
                tbodyTr.appendChild(td);
            });
            tbody.appendChild(tbodyTr);
        });
        table.appendChild(tbody);

        while (tableElement.firstChild) {
            tableElement.removeChild(tableElement.firstChild);
        }
        tableElement.appendChild(table);
    }

    loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/jszip.js', () => {
        loadScript('https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.8.0/xlsx.js', () => {
            if (options.autostart) {
                autostart(selector, options);
                if (queue.length) {
                    let queueData;
                    while (queueData = queue.pop()) {
                        processTable(queueData[0], queueData[1]);
                    }
                }
            }
        });
    });

    function fireEvent(event, options) {
        if (typeof options[event] === 'function') {
            options[event].call(this, options);
        }
    }

    function processTable(selector, options) {
        options = Object.assign({
            tableClass: 'xlsSortable-table',
            tableSortClass: 'xlsSortable-sortable',
            theme: false,
        }, options || {});

        if (typeof options.guessFunctions !== 'undefined') {
            options.guessFunctions.reverse().forEach(guessFunction => {
                if (typeof guessFunction === 'function') {
                    guessFunctions.unshift(guessFunction);
                }
            });
        }
        if (typeof options.sortFunctions !== 'undefined') {
            Object.assign(sortFunctions, options.sortFunctions);
        }
        if (typeof options.printFunctions !== 'undefined') {
            Object.assign(printFunctions, options.printFunctions);
        }

        if (typeof XLSX === 'undefined') {
            queue.push([selector, options]);
            return true;
        }
        if (typeof selector !== 'undefined') {
            if (typeof selector === 'string') {
                let elements = document.querySelectorAll(selector);
                if (elements) {
                    elements.forEach(element => {
                        processTable(element, options);
                    });
                }
            } else {
                let tableOptions = Object.assign(options, selector.dataset);
                if (typeof tableOptions.noSort === 'string') {
                    tableOptions.noSort = tableOptions.noSort.toUpperCase().split(',');
                }
                tableOptions.el = selector;
                loadTable(options.file, XLSobject => {
                    let tableData = {headers: [], body: [], types: [], minValues: [], maxValues: []},
                        typesCount = [];
                    if (XLSobject[0]) {
                        for (let header in XLSobject[0]) {
                            if (!XLSobject[0].hasOwnProperty(header)){
                                continue;
                            }
                            tableData.headers.push(header);
                            typesCount[header] = {};
                        }
                    }
                    XLSobject.forEach(dataRow => {
                        let row = [];
                        for (let i in dataRow) {
                            row.push(typeof dataRow[i] !== 'undefined' ? dataRow[i] : '');
                            let type = guessType(dataRow[i], i);
                            if (typeof typesCount[i][type] === 'undefined') {
                                typesCount[i][type] = 0;
                            }
                            ++typesCount[i][type];
                        }
                        tableData.body.push(row);
                    });
                    for (let i in typesCount) {
                        let col = tableData.types.length;
                        let colName = numToRow(col);
                        let optionType = 'type' + colName;
                        let optionMinValue = 'minValue' + colName;
                        let optionMaxValue = 'maxValue' + colName;
                        if (typeof options[optionType] !== 'undefined') {
                            tableData.types.push(options[optionType]);
                        } else {
                            let maxType = {
                                type: '',
                                count: 0,
                            };
                            for (let type in typesCount[i]) {
                                if (typesCount[i][type] > maxType.count) {
                                    maxType.type = type;
                                    maxType.count = typesCount[i][type];
                                }
                            }
                            tableData.types.push(maxType.type);
                        }
                        if (typeof options[optionMinValue] !== 'undefined') {
                            tableData.minValues[col] = options[optionMinValue];
                        }
                        if (typeof options[optionMaxValue] !== 'undefined') {
                            tableData.maxValues[col] = options[optionMaxValue];
                        }
                    }
                    tableData.body.forEach((dataRow, i) => {
                        dataRow.forEach((dataCell, j) => {
                            tableData.body[i][j] = printFunctions[tableData.types[j]](dataCell);
                        });
                    });
                    formTable(selector, tableData, tableOptions);
                    fireEvent('tableLoaded', tableOptions);
                });
            }
        }
    }

    function autostart(selector, options) {
        selector = selector || '.xls-table';
        options = options || {};
        processTable(selector, options);
    }

    return {
        processTable,
    };

};