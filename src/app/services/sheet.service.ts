import { Injectable } from '@angular/core';
import * as XLSX from 'xlsx';

@Injectable({
  providedIn: 'root'
})
export class SheetService {

   toExportFileName(sheetsFileName: string, fileFormat?: string): string {
    return `${sheetsFileName}${fileFormat ? fileFormat : '.xlsx'}`;
  }

  readFile(path: any, options?: object) {
    XLSX.read(path, options);
  }

  exportAsSheetsFile(json: any[], fileName:string, fileFormat?: string): void {
    const worksheet: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
    const workbook: XLSX.WorkBook = {
      Sheets: { data: worksheet },
      SheetNames: ['data']
    };
    return XLSX.writeFile(
      workbook,
      this.toExportFileName(fileName, fileFormat)
    );
  }
}
