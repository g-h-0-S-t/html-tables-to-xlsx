javascript:(() => {
    const loadSheetJS = () => new Promise((resolve, reject) => {
        if (typeof XLSX !== 'undefined') return resolve();
        const script = document.createElement('script');
        script.src = 'https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js';
        script.onload = resolve;
        script.onerror = reject;
        document.head.appendChild(script);
    });

    const getSheetData = () => {
        const sheetData = [];
        document.querySelectorAll('table').forEach(table => {
            table.querySelectorAll('tr').forEach(row => {
                const rowData = [];
                row.querySelectorAll('th, td').forEach(cell => rowData.push(cell.innerText));
                sheetData.push(rowData);
            });
            sheetData.push([]);
        });
        return XLSX.utils.aoa_to_sheet(sheetData);
    };

    const exportTablesToExcel = async () => {
        try {
            await loadSheetJS();
            const sheet = getSheetData();
            if (sheet.length === 0) return alert('No tables found on this page.');
            const workbook = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(workbook, sheet, 'Tables');
            XLSX.writeFile(workbook, 'exported_tables.xlsx');
        } catch (error) {
            console.error('Error exporting tables to Excel:', error);
        }
    };

    exportTablesToExcel();
})();
