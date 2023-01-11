import { Component, ElementRef, ViewChild } from '@angular/core';
import { IgxSpreadsheetComponent, SpreadsheetCell, SpreadsheetCellEditMode } from 'igniteui-angular-spreadsheet';
import { ExcelUtility } from './ExcelUtility';
import { FormatConditionAboveBelow, Workbook, WorkbookColorInfo, WorkbookSaveOptions } from 'igniteui-angular-excel';
import { Color } from 'igniteui-angular-core';
import { AppService } from './app.service';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
//import { Workbook } from "exceljs";

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.scss']
})
export class AppComponent {
  @ViewChild("spreadsheet", { read: IgxSpreadsheetComponent, static: true })
  public spreadsheet!: IgxSpreadsheetComponent;
  //@ViewChild('select') select! : HTMLSelectElement;
  @ViewChild('select') select!: ElementRef;
  title = 'LoadSpreadsheet';
  blue = new Color();
  empList: Array<any> = [];
  salesList: Array<any> = [];
  ExcelData: any;
  
  public isProtected: boolean;
  constructor(private appSerrvice : AppService,private httpClient: HttpClient) {
    this.isProtected = false; this.blue.colorString = "#ff0000";
    this.empList = [
      { empid: 101, name: "Rohit" },
      { empid: 102, name: "Mohit" },
      { empid: 103, name: "Jack" },
      { empid: 104, name: "Smith" }
    ];
    this.salesList = [
      { id: 1, name: "Soft Beverages" , empid: 101 },
      { id: 2, name: "Bottled Beer" , empid: 102 },
      { id: 3, name: "Draft Beer" , empid: 103 },
      { id: 4, name: "Liquor" , empid: 104 },
      { id: 5, name: "Wine" , empid: 105 },
      { id: 6, name: "Snacks" , empid: 106 },
      { id: 7, name: "Potato Chips" , empid: 107 },
      { id: 8, name: "Nuts" , empid: 108 }
    ]
  }
  ngOnInit() {
    const excelFile = '../../assets/sheets/ParentWorkbook.xlsx';
    ExcelUtility.loadFromUrl(excelFile).then((w) => {
      this.spreadsheet.workbook = w;
      
    });
  }
  public openFile(input: HTMLInputElement): void {
    if (input.files == null || input.files.length === 0) {
    return;
    }
    ExcelUtility.load(input.files[0]).then((w) => {
    this.spreadsheet.workbook = w;
    }, (e) => {
        console.error("Workbook Load Error:" + e);
    });
  }

  public workbookSave(): void {
    const opt = new WorkbookSaveOptions();
    opt.type = "blob";
    this.spreadsheet.workbook.save(opt, (d) => {
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(d as Blob);
      fileReader.onload = (e: any) => {
        var workbook = XLSX.read(fileReader.result, { type: 'binary' });
        var sheetNames = workbook.SheetNames;
        this.ExcelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
      }
    }, (e) => {
    });
  }
  public workbookDownload(): void {
    ExcelUtility.save(this.spreadsheet.workbook, ".xlsx");
  }
  public AddRowColumn() {
    // this.shortcuts.push({
    //   key: ["cmd" + "Shift" + "+"],
    //   allowIn: [AllowIn.Textarea, AllowIn.Input],
    //   command: e => console.error(`Ctrl+shift+P has been hijacked`),
    //   preventDefault: true,
    // });
    // let avgFormat = this.spreadsheet.activeWorksheet.conditionalFormats().addAverageCondition("B1:B10", FormatConditionAboveBelow.AboveAverage);
    // avgFormat.cellFormat.font.colorInfo = new WorkbookColorInfo(this.blue);
    // let uniqueFormat = this.spreadsheet.activeWorksheet.conditionalFormats().addUniqueCondition("O1:O10");
    // uniqueFormat.cellFormat.font.colorInfo = new WorkbookColorInfo(this.blue);
  }
  public onChange() {
    this.spreadsheet.activeWorksheet.protect();
    this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = false;
    if(this.select.nativeElement.value == "101") {
      this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    }
    if(this.select.nativeElement.value == "102") {
      this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    }
    if(this.select.nativeElement.value == "103") {
      this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    }
    if(this.select.nativeElement.value == "104") {
      this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
      this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = false;
      this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    }
    
  }
  public onProtectedChanged(e: any) {
    var as = e;
    if (e.target.checked) {
        this.spreadsheet.activeWorksheet.protect();
        this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = false;
    } else {
        this.spreadsheet.activeWorksheet.unprotect();
    }
  }
  
}
