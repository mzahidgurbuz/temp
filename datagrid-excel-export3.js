onExporting(e) {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Export");

    const grid = e.component;

    const masterColumns = grid.getVisibleColumns()
        .filter(c => c.dataField)
        .map(c => ({
            field: c.dataField,
            caption: c.caption,
            lookup: c.lookup
        }));


    const detailGrid = $("#detailGrid").dxDataGrid("instance");

    const detailColumns = detailGrid.getVisibleColumns()
        .filter(c => c.dataField)
        .map(c => ({
            field: c.dataField,
            caption: c.caption,
            lookup: c.lookup
        }));


    const headers = [
        ...masterColumns.map(c => c.caption),
        ...detailColumns.map(c => c.caption)
    ];

    worksheet.addRow(headers);


    const data = grid.getDataSource().items();

    data.forEach(user => {

        if (user.logs && user.logs.length > 0) {

            user.logs.forEach(log => {

                const row = [];

                masterColumns.forEach(col => {

                    let value = user[col.field];

                    // lookup varsa display değeri yaz
                    if (col.lookup) {

                        const ds = col.lookup.dataSource;
                        const valueExpr = col.lookup.valueExpr;
                        const displayExpr = col.lookup.displayExpr;

                        const item = ds.find(x => x[valueExpr] === value);

                        if (item) value = item[displayExpr];
                    }

                    row.push(value);
                });


                detailColumns.forEach(col => {

                    let value = log[col.field];

                    if (col.lookup) {

                        const ds = col.lookup.dataSource;
                        const valueExpr = col.lookup.valueExpr;
                        const displayExpr = col.lookup.displayExpr;

                        const item = ds.find(x => x[valueExpr] === value);

                        if (item) value = item[displayExpr];
                    }

                    row.push(value);
                });

                const excelRow = worksheet.addRow(row);

                excelRow.eachCell((cell) => {

                    if (typeof cell.value === "number") {

                        const str = cell.value.toString();

                        if (str.length > 15) {

                            cell.value = str;
                            cell.numFmt = "@";
                        }
                    }

                });

            });

        }

    });


    workbook.xlsx.writeBuffer().then(buffer => {
        saveAs(new Blob([buffer]), "Export.xlsx");
    });

    e.cancel = true;
}
