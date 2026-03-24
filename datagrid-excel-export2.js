function exportGridAdvanced(e, fileName = "export.xlsx") {

    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet("Data");

    DevExpress.excelExporter.exportDataGrid({
        component: e.component,
        worksheet: worksheet,
        autoFilterEnabled: true,

        customizeCell(options) {

            const gridCell = options.gridCell;
            const excelCell = options.excelCell;

            if (!gridCell) return;

            let value = gridCell.value;

            // null fix
            if (value === null || value === undefined) {
                excelCell.value = "";
                return;
            }

            // lookup kolonları display text
            if (gridCell.column.lookup && gridCell.displayValue) {
                excelCell.value = gridCell.displayValue;
                return;
            }

            // nested object flatten
            if (typeof value === "object") {
                excelCell.value = JSON.stringify(value);
                return;
            }

            // Excel 15 digit fix
            if (typeof value === "number" && value.toString().length > 15) {
                excelCell.value = value.toString();
                excelCell.numFmt = "@";
                return;
            }

            // bigint string fix
            if (typeof value === "string" && /^\d{15,}$/.test(value)) {
                excelCell.numFmt = "@";
                return;
            }

        }

    }).then(() => {

        // kolon başlıklarını otomatik bold yap
        worksheet.getRow(1).font = { bold: true };

        workbook.xlsx.writeBuffer().then(buffer => {

            saveAs(
                new Blob([buffer], { type: "application/octet-stream" }),
                fileName
            );

        });

    });

    e.cancel = true;
}
