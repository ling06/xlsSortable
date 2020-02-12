var XlsReader = function (selector, options) {
    options = Object.assign({
        autostart: true,
    }, options || {});

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

    function sortData(data, sortBy, sortDir) {
        data = data.sort((a, b) => {
            return (sortDir === 'asc' && a[sortBy] > b[sortBy]) || (sortDir === 'desc' && a[sortBy] < b[sortBy]) ?
                1 :
                (
                    (sortDir === 'asc' && a[sortBy] < b[sortBy]) || (sortDir === 'desc' && a[sortBy] > b[sortBy]) ?
                        -1 :
                        0
                );
        });
        return data;
    }

    function formTable(tableElement, tableData, options) {
        options = Object.assign({
            sortBy: -1,
            sortDir: '',
        }, options || {});
        if (options.sortBy !== -1 && options.sortDir) {
            tableData.body = sortData(tableData.body, options.sortBy, options.sortDir);
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
            th.classList.add(options.tableSortClass);
            if (options.sortBy === i) {
                th.classList.add(options.tableSortClass + '-' + options.sortDir);
            }
            th.innerText = tableData.headers[i];
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
                    sortData(tableData.body, options.sortBy, options.sortDir);
                    formTable(tableElement, tableData, options);
                });
            })(i);
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

    function processTable(selector, options) {
        options = Object.assign({
            tableClass: 'xlsReader-table',
            tableSortClass: 'xlsReader-sortable',
            theme: false,
        }, options || {});

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
                loadTable(options.file, XLSobject => {
                    let tableData = {headers: [], body: [], types: []},
                        typesCount = [];
                    if (XLSobject[0]) {
                        for (let header in XLSobject[0]) {
                            if (!XLSobject[0].hasOwnProperty(header)){
                                continue;
                            }
                            tableData.headers.push(header);
                            typesCount[header] = {number: 0, string: 0};
                        }
                    }
                    XLSobject.forEach(dataRow => {
                        let row = [];
                        for (let i in dataRow) {
                            row.push(typeof dataRow[i] !== 'undefined' ? dataRow[i] : '');
                            if (isNaN(parseFloat(dataRow[i]))) {
                                ++typesCount[i].string;
                            } else {
                                ++typesCount[i].number;
                            }
                        }
                        tableData.body.push(row);
                    });
                    for (let i in typesCount) {
                        tableData.types.push(typesCount[i].number >= typesCount[i].string ? 'number' : 'string');
                    }
                    tableData.body.forEach((dataRow, i) => {
                        dataRow.forEach((dataCell, j) => {
                            if (tableData.types[j] === 'number') {
                                tableData.body[i][j] = parseFloat(dataCell);
                            }
                        });
                    });
                    formTable(selector, tableData, tableOptions);
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