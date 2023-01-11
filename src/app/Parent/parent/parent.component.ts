import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { IgxSpreadsheetComponent, SpreadsheetCell, SpreadsheetCellEditMode } from 'igniteui-angular-spreadsheet';
import { FormatConditionAboveBelow, Workbook, WorkbookColorInfo, WorkbookSaveOptions, WorksheetHyperlink } from 'igniteui-angular-excel';
import { Color } from 'igniteui-angular-core';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
import { AppService } from 'src/app/app.service';
import { ExcelUtility } from 'src/app/ExcelUtility';
//import { Workbook } from "exceljs";

@Component({
  selector: 'app-parent',
  templateUrl: './parent.component.html',
  styleUrls: ['./parent.component.scss']
})
export class ParentComponent implements OnInit {

  @ViewChild("spreadsheet", { read: IgxSpreadsheetComponent, static: true })
  public spreadsheet!: IgxSpreadsheetComponent;
  //@ViewChild('select') select! : HTMLSelectElement;
  @ViewChild('select') select!: ElementRef;
  title = 'LoadSpreadsheet';
  blue = new Color();
  empList: Array<any> = [];
  salesList: Array<any> = [];
  ExcelData: any;
  parentData: any;
  excelFile: any;
  loadThis: boolean = false;
  existOrNot: boolean = false;

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

  ngDoCheck() {
    if (this.loadThis)
      this.appSerrvice.getParentWorbookData().subscribe(res => {
        this.parentData = res;
        this.parentData.forEach(function (v: any) { delete v.product });
        const opt = new WorkbookSaveOptions();
        opt.type = "blob";
        this.spreadsheet.workbook.save(opt, (d) => {
          let fileReader = new FileReader();
          fileReader.readAsBinaryString(d as Blob);
          fileReader.onload = (e: any) => {
            var workbook = XLSX.read(fileReader.result, { type: 'binary' });
            var sheetNames = workbook.SheetNames;
            this.ExcelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
            var comparData = this.ExcelData.slice(0, this.ExcelData.length-1);
            this.parentData.forEach((o: { [x: string]: any; }) => {
              delete o['products'];
            });
            const sheet1 : any= {};
            var count = 0;
            comparData.forEach((x:any)=>{
                 var values = { name : x.Name , monday : x.Monday , tuesday : x.Tuesday , 
                  wednesday : x.Wednesday, thursday : x.Thursday, friday : x.Friday, saturday : x.Saturday, sunday : x.Sunday };
                  sheet1[count] = values;
                  count ++;
            })
            const sheet2 : any= {};
            count = 0;
            this.parentData.forEach((x:any)=>{
                 var values = { name : x.name , monday : x.monday , tuesday : x.tuesday , 
                  wednesday : x.wednesday, thursday : x.thursday, friday : x.friday, saturday : x.saturday, sunday : x.sunday };
                  sheet2[count] = values;
                  count ++;
            })
            const comparDataCount = comparData.length;
            const ExcelDataCount = this.parentData.length;
            if(comparDataCount === ExcelDataCount) {
              this.existOrNot = (JSON.stringify(sheet1) === JSON.stringify(sheet2)) 
            } else {
              this.existOrNot = false;
            }
            if (!this.existOrNot) { 
              window.location.reload();
              this.loadThis = true;
            } else {
              this.loadThis = true;
            }
          }
        }, (e) => {
        });
      });
  }
  
  ngOnInit() {
    this.appSerrvice.getFile().subscribe(data => {
      var file = new File([data], "parent");
      ExcelUtility.load(file).then((w) => {
        this.appSerrvice.getParentWorbookData().subscribe(res => {
          this.parentData = res;
          var c = 1;
          this.spreadsheet.workbook = w;
          this.parentData.forEach(function (v: any) { delete v.product });
          this.parentData.forEach((x: any) => {
            this.spreadsheet.activeWorksheet.rows(c).cells(0).value = x.name;
            this.spreadsheet.activeWorksheet.rows(c).cells(1).value = x.monday;
            this.spreadsheet.activeWorksheet.rows(c).cells(2).value = x.tuesday;
            this.spreadsheet.activeWorksheet.rows(c).cells(3).value = x.wednesday;
            this.spreadsheet.activeWorksheet.rows(c).cells(4).value = x.thursday;
            this.spreadsheet.activeWorksheet.rows(c).cells(5).value = x.friday;
            this.spreadsheet.activeWorksheet.rows(c).cells(6).value = x.saturday;
            this.spreadsheet.activeWorksheet.rows(c).cells(7).value = x.sunday;
            c++;
          });
          this.loadThis = true;
        });
      });
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
    this.workbookDataSave();
  }
  public workbookDataSave(){
    const opt = new WorkbookSaveOptions();
    opt.type = "blob";
    this.spreadsheet.workbook.save(opt, (d) => {
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(d as Blob);
      // var excelFile = '../../assets/sheets/ParentWorkbook.xlsx';
      const formData = new FormData();
      formData.append('file', d as Blob, "ParentWorkbookData.xlsx");
      this.appSerrvice.getFileUpload(formData).subscribe(res => {
       
      });
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
