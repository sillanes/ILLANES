
document.writeln('<script lang="javascript" src="../includes/excel_lib/xlsx_sheetjs.js" ></script>');
document.writeln('<script lang="javascript" src="../includes/excel_lib/file_saver.js" ></script>');


	function s2ab(data) {
        const buf = new ArrayBuffer(data.length);
        const view = new Uint8Array(buf);
        for (let i=0; i < data.length; i++) view[i] = data.charCodeAt(i) & 0xFF;
        return buf;
    }
	
	
	function exportTableToXlsx(file_name, sheets = []) {
        const book = XLSX.utils.book_new()

        sheets.forEach((item) => {			
            const wb = XLSX.utils.table_to_sheet(document.getElementById(item), {raw:true});
            book.SheetNames.push(item)
            book.Sheets[item] = wb
        })
		
		const today = new Date();
		const dd = String(today.getDate()).padStart(2, '0');
		const mm = String(today.getMonth() + 1).padStart(2, '0');
		const yyyy = today.getFullYear();

		current_date =  yyyy + mm + dd;

       
        const wbout = XLSX.write(book, { bookType:'xlsx', bookSST:true, type: 'binary' });
		saveAs(new Blob([s2ab(wbout)], { type:"application/octet-stream" }), `${file_name}_${current_date}.xlsx`);
    }
	
	
	function exportTableToXlsx2(params) {		
        const book = XLSX.utils.book_new()

		const { name, ids, data } = params;
		
		let doc = new DOMParser().parseFromString(data, 'text/html');
		//let doc = document.createRange().createContextualFragment(data);
		
		ids.forEach((id, index) => {			
            const wb = XLSX.utils.table_to_sheet(doc.body.childNodes[index], {raw:true});
            book.SheetNames.push(id)
            book.Sheets[id] = wb
        })		
        		
		const today = new Date();
		const dd = String(today.getDate()).padStart(2, '0');
		const mm = String(today.getMonth() + 1).padStart(2, '0');
		const yyyy = today.getFullYear();

		current_date =  yyyy + mm + dd;
		doc = null; 
		
        const wbout = XLSX.write(book, { bookType:'xlsx', bookSST:true, type: 'binary' });
		saveAs(new Blob([s2ab(wbout)], { type:"application/octet-stream" }), `${name}_${current_date}.xlsx`);
    }
