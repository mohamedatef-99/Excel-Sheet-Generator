let table = document.getElementsByClassName("sheet-body")[0],
    rows = document.getElementsByClassName("rows")[0],
    columns = document.getElementsByClassName("columns")[0]
    tableExists = false

    const generateTable = () => {
        let rowsNumber = parseInt(rows.value), columnsNumber = parseInt(columns.value);
        table.innerHTML = "";
        
        if (isNaN(rowsNumber) || isNaN(columnsNumber) || rowsNumber <= 0 || columnsNumber <= 0) {
            Swal.fire({
                icon: 'error',
                title: 'Oops...',
                text: 'Please enter valid rows and columns!'
            });
            return;
        }
        for(let i=0; i<rowsNumber; i++){
            var tableRow = "";
            for(let j=0; j<columnsNumber; j++){
                tableRow += `<td contenteditable></td>`;
            }
            table.innerHTML += tableRow;
        }
        tableExists = true;
    }

const ExportToExcel = (type, fn, dl) => {
    if (!tableExists || !table.querySelector('td')) {
        Swal.fire({
            icon: 'error',
            title: 'Oops...',
            text: 'Table is empty! Please add some data to export.'
        });
        return;
    }
    var elt = table
    var wb = XLSX.utils.table_to_book(elt, { sheet: "sheet1" })
    return dl ? XLSX.write(wb, { bookType: type, bookSST: true, type: 'base64' })
        : XLSX.writeFile(wb, fn || ('MyNewSheet.' + (type || 'xlsx')))
}
