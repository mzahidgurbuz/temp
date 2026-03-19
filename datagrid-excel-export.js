onExporting: async function(e) {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Users");

    await DevExpress.excelExporter.exportDataGrid({
        component: e.component,
        worksheet: worksheet,
        autoFilterEnabled: true
    });

    const detailGrids = e.component.getView("rowsView")._rows
        .filter(r => r.rowType === "detail");

    for (const row of detailGrids) {

        const detailGrid = $(row.rowElement)
            .find(".dx-datagrid")
            .dxDataGrid("instance");

        await DevExpress.excelExporter.exportDataGrid({
            component: detailGrid,
            worksheet: worksheet,
            topLeftCell: { row: worksheet.rowCount + 1, column: 2 }
        });

    }

    workbook.xlsx.writeBuffer().then(function(buffer) {
        saveAs(new Blob([buffer]), "Users.xlsx");
    });

    e.cancel = true;
}