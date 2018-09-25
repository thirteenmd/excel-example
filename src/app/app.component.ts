import { Component } from '@angular/core';

import * as XLSX from 'xlsx';

type AOA = any[][];

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  fixedRows: AOA = [[], []];
  values: AOA = [[], []];

  wopts: XLSX.WritingOptions = { bookType: 'xlsx', type: 'array' };
  fileName: string = 'ExportedFile.xlsx';

  onFileChange(evt: any) {
    /* wire up file reader */
    let data = [[], []];

    const target: DataTransfer = <DataTransfer>(evt.target);
    if (target.files.length !== 1) throw new Error('Cannot use multiple files');
    const reader: FileReader = new FileReader();
    reader.onload = (e: any) => {
      /* read workbook */
      const bstr: string = e.target.result;
      const wb: XLSX.WorkBook = XLSX.read(bstr, { type: 'binary' });

      /* grab first sheet */
      const wsname: string = wb.SheetNames[0];
      const ws: XLSX.WorkSheet = wb.Sheets[wsname];

      /* save data */
      data = <AOA>(XLSX.utils.sheet_to_json(ws, { header: 1 }));
      this.fixedRows = data.slice(0, 3);
      this.values = data.slice(3);
    };
    reader.readAsBinaryString(target.files[0]);

  }

  export(): void {
    let data = this.fixedRows.concat(this.values);
    /* generate worksheet */
    const ws: XLSX.WorkSheet = XLSX.utils.aoa_to_sheet(data);

    /* generate workbook and add the worksheet */
    const wb: XLSX.WorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

    /* save to file */
    XLSX.writeFile(wb, this.fileName);
  }

  changeValue(i: number, j: number, event: any) {
    const newValue = event.target.textContent;
    this.values[i][j] = newValue;
  }

}
