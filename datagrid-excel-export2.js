function exportMasterDetailGrid(gridSelector, detailGridSelector, fileName = "Export.xlsx") {

    const grid = $(gridSelector).dxDataGrid("instance");
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Export");

    // Master kolonları
    const masterColumns = grid.getVisibleColumns()
        .filter(c => c.dataField)
        .map(c => ({
            field: c.dataField,
            caption: c.caption,
            lookup: c.lookup
        }));

    // Detail kolonları (template grid)
    const detailGrid = $(detailGridSelector).dxDataGrid("instance");
    const detailColumns = detailGrid.getVisibleColumns()
        .filter(c => c.dataField)
        .map(c => ({
            field: c.dataField,
            caption: c.caption,
            lookup: c.lookup
        }));

    // Header üret
    const headers = [
        ...masterColumns.map(c => c.caption),
        ...detailColumns.map(c => c.caption)
    ];
    worksheet.addRow(headers);

    // Grid data
    const data = grid.getDataSource().items();

    data.forEach(user => {

        const logs = user.logs && user.logs.length > 0 ? user.logs : [{}];

        logs.forEach(log => {
            const row = [];

            // Master kolon değerleri
            masterColumns.forEach(col => {
                let value = user[col.field];

                if (col.lookup && col.lookup.dataSource) {
                    const ds = col.lookup.dataSource;
                    const item = ds.find(x => x[col.lookup.valueExpr] === value);
                    if (item) value = item[col.lookup.displayExpr];
                }

                row.push(value);
            });

            // Detail kolon değerleri
            detailColumns.forEach(col => {
                let value = log[col.field];

                if (col.lookup && col.lookup.dataSource) {
                    const ds = col.lookup.dataSource;
                    const item = ds.find(x => x[col.lookup.valueExpr] === value);
                    if (item) value = item[col.lookup.displayExpr];
                }

                row.push(value);
            });

            const excelRow = worksheet.addRow(row);

            // 15+ haneli sayıları string olarak yaz
            excelRow.eachCell(cell => {
                if (typeof cell.value === "number") {
                    const str = cell.value.toString();
                    if (str.length > 15) {
                        cell.value = str;
                        cell.numFmt = "@";
                    }
                }
            });

        });

    });

    // Excel kaydet
    workbook.xlsx.writeBuffer().then(buffer => {
        saveAs(new Blob([buffer]), fileName);
    });

}
