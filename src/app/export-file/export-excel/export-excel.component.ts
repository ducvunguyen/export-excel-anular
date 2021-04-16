import { Component, OnInit } from '@angular/core';
import * as logoFile from '../../carlogo';
import {Workbook} from 'exceljs';
import * as fs from 'file-saver';
import {POSITION_CELLS} from '../constants';
import {element} from 'protractor';
@Component({
  selector: 'app-export-excel',
  templateUrl: './export-excel.component.html',
  styleUrls: ['./export-excel.component.css']
})
export class ExportExcelComponent implements OnInit {
  dataSource = [
    [1, 'Khối 1', 3, 4, 4, 4, 4, 4, 4, 4, 10, 10],
    [2, 'Khối 2', 3, 4, 4, 4, 4, 4, 4, 4, 10, 10],
    [3, 'Khối 3', 3, 4, 4, 4, 4, 4, 4, 4, 10, 10],
    [4, 'Khối 4', 3, 4, 4, 4, 4, 4, 4, 4, 10, 10],
  ]

  constructor() {
  }

  ngOnInit() {

  }

  clickExport(){
    this.generateExcel();
  }
  async generateExcel() {
    let titleName = 'Trường THPT OMT PHÒNG GD Đống Đa';
    let labelSubject = ['' ,'' ,'Toán', 'Văn', 'Anh ', 'Sinh', 'Sử', 'Địa', 'Lý', 'Hóa', 'GDCD', 'Kỹ năng sống'];
    let borderGeneral = {
      top: {style:'thin', color: {argb:'000000'}},
      left: {style:'thin', color: {argb:'000000'}},
      bottom: {style:'thin', color: {argb:'000000'}},
      right: {style:'thin', color: {argb:'000000'}}
    };

    let fontGeneralHeader = {
      name: 'Calibri',
      color: { argb: '000000' },
      family: 4,
      size: 11,
      bold: true,
    };

    let fillBgHeader = {
      type: 'pattern',
      pattern: 'solid',
      fgColor: {argb: '2EFE9A'},
      bgColor:{argb:'4C0B5F'}
    };

    const workbook = new Workbook();
    const worksheet = workbook.addWorksheet('Thống kê lớp học');

    // Add Row and formatting
    const titleRow = worksheet.addRow([titleName]);
    titleRow.font = {name: 'Calibri', family: 4, size: 11, underline: 'none', bold: false};
    let verticalCenter = { vertical: 'middle', horizontal: 'center',  wrapText: true };
    worksheet.mergeCells('A1:B2');
    worksheet.mergeCells(`C1:${POSITION_CELLS[labelSubject.length-1]}1`);
    worksheet.mergeCells(`C2:${POSITION_CELLS[labelSubject.length-1]}2`);

    worksheet.getCell('C1').value = 'Thống kê lớp học'.toUpperCase();
    worksheet.getCell('C1').alignment = verticalCenter;
    worksheet.getCell('C1').font = {name: 'Calibri', size: 20, bold: true };
    worksheet.getCell('A1').alignment = verticalCenter;

    worksheet.getCell('C2').value = 'Năm học: 2020 - 2021';
    worksheet.getCell('C2').alignment = verticalCenter;
    worksheet.getRow(1).height = 30;

    labelSubject.forEach((item, index) => index == 0 ? null : worksheet.getColumn(index+1).width = 20);

    worksheet.getCell('A4').value = 'STT';
    worksheet.getCell('B4').value = 'Khối';
    worksheet.getCell('C4').value = 'Lớp';

    worksheet.getCell('A4').fill = fillBgHeader;
    worksheet.getCell('A4').font = fontGeneralHeader;
    worksheet.getCell('A4').border = borderGeneral;

    worksheet.getCell('B4').fill = fillBgHeader;
    worksheet.getCell('B4').font = fontGeneralHeader;

    worksheet.getCell('B4').border = borderGeneral;
    worksheet.getCell('B4').font = fontGeneralHeader;

    worksheet.getCell('C4').fill = fillBgHeader;
    worksheet.getCell('C4').font = fontGeneralHeader;
    worksheet.getCell('C4').border = borderGeneral

    const headerRowSubject = worksheet.addRow(labelSubject);
    headerRowSubject.eachCell((cell, number) => {
      cell.fill = fillBgHeader;
      cell.border = borderGeneral;
      cell.alignment = verticalCenter;
      cell.font = fontGeneralHeader;
    });

    worksheet.mergeCells('A4:A5');
    worksheet.mergeCells('B4:B5');
    worksheet.mergeCells(`C4:${POSITION_CELLS[labelSubject.length-1]}4`);
    worksheet.getCell('A4').alignment = verticalCenter;
    worksheet.getCell('B4').alignment = verticalCenter;
    worksheet.getCell('C4').alignment = verticalCenter;

    this.dataSource.forEach(element => {
      let addRow = worksheet.addRow(element);
      addRow.eachCell((cell, index) => {
        index !== 2 ? cell.alignment = verticalCenter : null;
        cell.border = {
          top: {style:'thin', color: {argb:'000000'}},
          left: {style:'thin', color: {argb:'000000'}},
          bottom: {style:'thin', color: {argb:'000000'}},
          right: {style:'thin', color: {argb:'000000'}}
        };
      })
    });

    worksheet.addRow([]);
    worksheet.addRow(['Lưu ý: Danh sách lớp học được triển khai trong năm học bao gồm các lớp đang dạy, đã kết thúc; không bao gồm các lớp học đã hủy, chưa triển khai']);

    workbook.xlsx.writeBuffer().then((data: any) => {
      const blob = new Blob([data], {type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
      fs.saveAs(blob, 'thong_ke_lop_hoc.xlsx');
    });

  }
}
