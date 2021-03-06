import {Injectable} from '@angular/core';
import {Workbook} from 'exceljs';
import * as fs from 'file-saver';
import * as logoFile from './carlogo.js';

// import { DatePipe } from '@angular/common';

@Injectable({
  providedIn: 'root'
})
export class ExcelService {

  constructor() {

  }

  async generateExcel(title: string, header: any[], data: any, nameFile: string, colSpan: number[]) {
    // const ExcelJS = await import('exceljs');
    // console.log(ExcelJS);
    // const Workbook: any = {};


    // Create workbook and worksheet
    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Car Data');


    // Add Row and formatting
    const titleRow = worksheet.addRow([title]);
    titleRow.font = {name: 'Comic Sans MS', family: 4, size: 16, underline: 'double', bold: true};
    worksheet.addRow([]);
    // const subTitleRow = worksheet.addRow(['Date : ' + new Date()]);
    worksheet.addRow(['Date : ' + new Date()]);


    // Add Image
    const logo = workbook.addImage({
      base64: logoFile.logoBase64,
      extension: 'png',
    });

    worksheet.addImage(logo, 'E1:F3');
    worksheet.mergeCells('A1:D2');

    worksheet.mergeCells('A6:A9')
    worksheet.mergeCells('A10:D10')
    // Blank Row
    worksheet.addRow([]);

    // Add Header Row
    const headerRow = worksheet.addRow(header);

    // Cell Style : Fill and Border
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: 'pattern',
        pattern: 'solid',
        fgColor: {argb: 'FF8038'},
        bgColor:{argb:'4C0B5F'}
      };
      cell.border = {top: {style: 'dotted'}, left: {style: 'dotted'}, bottom: {style: 'dotted'}, right: {style: 'dotted'}};
    });
    // worksheet.addRows(data);


    // Add Data and Conditional Formatting
    // let row = 0;
    // test.forEach(item => {
    //   data.forEach(d => {
    //     const row = worksheet.addRow(d);
    //     const qty = row.getCell(1);
    //     // if (+qty.value == item)worksheet.mergeCells(`A${item}:A${item}`)
    //   });
    // })

    data.forEach(d => {
        const row = worksheet.addRow(d);
        const qty = row.getCell(6); //chon colums de xu ly
        let color = 'FF99FF99';
        if (+qty.value < 500) {
          color = '3104B4';
        }

        qty.fill = {
          type: 'pattern',
          pattern: 'solid',
          fgColor: {argb: color}
        };
      }
    );

    worksheet.getColumn(1).width = 30;//tang kich thuoc cua cot
    worksheet.getColumn(3).width = 30;
    worksheet.getColumn(4).width = 30;
    worksheet.addRow([]);

    let customTest = [...new Set(colSpan)];
    let dataTest = [];
    let row = 13;
    customTest.forEach(item => dataTest.push(colSpan.filter(col => col === item))) ;
    dataTest.forEach(item =>{
      worksheet.mergeCells(`A${row}:A${row + item.length - 1}`);
      row += item.length;
    });


    // Footer Row
    const footerRow = worksheet.addRow(['This is system generated excel sheet.']);
    footerRow.getCell(1).fill = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: 'FFCCFFE5'}
    };
    footerRow.getCell(1).border = {top: {style: 'thin'}, left: {style: 'thin'}, bottom: {style: 'thin'}, right: {style: 'thin'}};

    // Merge Cells
    worksheet.mergeCells(`A${footerRow.number}:F${footerRow.number}`);
    worksheet.mergeCells('G1:H1');//merge 2 o voi nhau

    // Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      fs.saveAs(blob, nameFile+'.xlsx');
    });

  }
}
