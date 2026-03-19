 function onExporting(e) {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Users");

    DevExpress.excelExporter.exportDataGrid({
        component: e.component,
        worksheet: worksheet
    }).then(function () {

        const detailGrids = [];

        $(".dx-master-detail-row .dx-datagrid").each(function () {
            const grid = $(this).dxDataGrid("instance");
            if (grid) {
                detailGrids.push(grid);
            }
        });

        let chain = Promise.resolve();

        detailGrids.forEach(function(detailGrid) {
            chain = chain.then(function () {
                return DevExpress.excelExporter.exportDataGrid({
                    component: detailGrid,
                    worksheet: worksheet,
                    topLeftCell: { row: worksheet.rowCount + 1, column: 2 }
                });
            });
        });

        return chain;
    })
    .then(function () {
        return workbook.xlsx.writeBuffer();
    })
    .then(function (buffer) {
        saveAs(new Blob([buffer]), "Users.xlsx");
    });
    e.cancel = true;
}