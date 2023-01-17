import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { IgxSpreadsheetComponent, SpreadsheetCell, SpreadsheetCellEditMode } from 'igniteui-angular-spreadsheet';
import { CellReferenceMode, FormatConditionAboveBelow, Formula, Workbook, WorkbookColorInfo, WorkbookSaveOptions, WorksheetHyperlink } from 'igniteui-angular-excel';
import { Color } from 'igniteui-angular-core';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
import { AppService } from 'src/app/app.service';
import { AlphaBetica, ExcelUtility } from 'src/app/ExcelUtility';
//import { Workbook } from "exceljs";

@Component({
  selector: 'app-child1',
  templateUrl: './child1.component.html',
  styleUrls: ['./child1.component.scss']
})
export class Child1Component implements OnInit {

  @ViewChild("spreadsheet", { read: IgxSpreadsheetComponent, static: true })
  public spreadsheet!: IgxSpreadsheetComponent;
  //@ViewChild('select') select! : HTMLSelectElement;
  @ViewChild('select') select!: ElementRef;
  title = 'LoadSpreadsheet';
  blue = new Color();
  empList: Array<any> = [];
  salesList: Array<any> = [];
  ExcelData: any;
  child1Data: any;
  
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
    this.appSerrvice.getChild1File().subscribe(data => {
      var file = new File([data], "child");
      ExcelUtility.load(file).then((w) => {
        this.spreadsheet.workbook = w;
        
        //this.spreadsheet.activeWorksheet.sortSettings
        //this.spreadsheet.activeWorksheet.rows(1).cellFormat
        //this.spreadsheet.activeWorksheet.columns(2).cellFormat.setFormatting(this.spreadsheet.activeWorksheet.columns(1).cellFormat);
        //this.spreadsheet.activeWorksheet.rows(4).cells(1).applyFormula("=Sum(B2,B3,B4");

      });
    });
    // this.appSerrvice.getChild1WorbookData().subscribe(res => {
    //   var excelFile = '../../assets/sheets/Child1Workbook.xlsx';
    //   const fileName = 'test.xlsx';
    //   this.child1Data = res;
    //   this.child1Data.forEach(function(v: any){ delete v.name });
    //   // const ws: XLSX.WorkSheet = XLSX.utils.json_to_sheet(this.parentData);
    //   // const wb: XLSX.WorkBook = XLSX.utils.book_new();
    //   // XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
    //   //XLSX.writeFile(wb, fileName);
    //   this.loadChildExcelData(excelFile);
    // });
  }
  // public loadChildExcelData(excelFile: string ){
  //   ExcelUtility.loadFromUrl(excelFile).then((w) => {
  //     this.spreadsheet.workbook = w;
  //     var c = 1;
  //    this.child1Data.forEach((x:any)=> {
  //     this.spreadsheet.activeWorksheet.rows(c).cells(0).value = x.products;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(1).value = x.monday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(2).value = x.tuesday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(3).value = x.wednesday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(4).value = x.thursday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(5).value = x.friday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(6).value = x.saturday;
  //     this.spreadsheet.activeWorksheet.rows(c).cells(7).value = x.sunday;
  //     c++;
  //    });
  //   });
  // }
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
        //this.ExcelData.forEach(function(v: any){ delete v.total });
        // for(var j = 2; j <= this.ExcelData.length; j++){
        //   this.spreadsheet.activeWorksheet.hyperlinks().add(new WorksheetHyperlink("A" + j, "http://www.infragistics.com", "+", "Add Row Above"));
        // }
        // for (var i = 1; i <= 7; i++) {
        //   var sumFormula = Formula.parse("=SUM(" + AlphaBetica[i] + "1:" + AlphaBetica[i] + "" + this.ExcelData.length + ")", CellReferenceMode.A1);
        //   sumFormula.applyTo(this.spreadsheet.activeWorksheet.rows(this.ExcelData.length).cells(i));
        // }
        this.appSerrvice.InsertDealerDetails("child1", JSON.stringify(this.ExcelData)).subscribe((response: any) => {
          this.workbookSaveInFolder();
        });
      }
    }, (e) => {
    });
  }
  public workbookSaveInFolder(): void {
    const opt = new WorkbookSaveOptions();
    opt.type = "blob";
    this.spreadsheet.workbook.save(opt, (d) => {
      const formData = new FormData();
      formData.append('file', d as Blob, "Child1WorkbookData.xlsx");
      this.appSerrvice.getFileUpload(formData).subscribe(res => {
        alert("Inserted Records")
      });
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
