import { Component, ElementRef, OnInit, ViewChild } from '@angular/core';
import { IgxSpreadsheetComponent, SpreadsheetCell, SpreadsheetCellEditMode } from 'igniteui-angular-spreadsheet';
import { CellReferenceMode, FormatConditionAboveBelow, Formula, Workbook, WorkbookColorInfo, WorkbookSaveOptions } from 'igniteui-angular-excel';
import { Color } from 'igniteui-angular-core';
import * as XLSX from 'xlsx';
import { HttpClient } from '@angular/common/http';
import { AppService } from 'src/app/app.service';
import { AlphaBetica, ExcelUtility } from 'src/app/ExcelUtility';

@Component({
  selector: 'app-child2',
  templateUrl: './child2.component.html',
  styleUrls: ['./child2.component.scss']
})
export class Child2Component implements OnInit {

  @ViewChild("spreadsheet", { read: IgxSpreadsheetComponent, static: true })
  public spreadsheet!: IgxSpreadsheetComponent;
  //@ViewChild('select') select! : HTMLSelectElement;
  @ViewChild('select') select!: ElementRef;
  title = 'LoadSpreadsheet';
  blue = new Color();
  empList: Array<any> = [];
  salesList: Array<any> = [];
  ExcelData: any;
  child2Data: any;
  totalRowId: any;
  totalCount: any;
  
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
    this.appSerrvice.getChild2File().subscribe(data => {
      var file = new File([data], "child");
      ExcelUtility.load(file).then((w) => {
        this.spreadsheet.workbook = w;
        this.getTotalCount();
      });
    });
  }
  
  public openFile(input : any): void {

    const target: DataTransfer = <DataTransfer>(input.target);

    if (target.files == null || target.files.length === 0) {
    return;
    }
    ExcelUtility.load(target.files[0]).then((w) => {
    this.spreadsheet.workbook = w;
    }, (e) => {
        console.error("Workbook Load Error:" + e);
    });
  }
  getTotalCount(){
    const opt = new WorkbookSaveOptions();
    opt.type = "blob";
    this.spreadsheet.workbook.save(opt, (d) => {
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(d as Blob);
      fileReader.onload = (e: any) => {
        var workbook = XLSX.read(fileReader.result, { type: 'binary' });
        var sheetNames = workbook.SheetNames;
        this.ExcelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
        this.totalCount = this.ExcelData.length;
        this.totalRowId = "A" + Number(this.ExcelData.length + 1);
      }
    }, (e) => {
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
        if(this.ExcelData != null){
          for (var i = 1; i < Object.keys(this.ExcelData[0]).length; i++) {
            var sumFormula = Formula.parse("=SUM(" + AlphaBetica[i] + "1:" + AlphaBetica[i] + "" + this.ExcelData.length + ")", CellReferenceMode.A1);
            sumFormula.applyTo(this.spreadsheet.activeWorksheet.rows(this.ExcelData.length).cells(i));
          }
        }
        this.workbookSaveData();
      }
    }, (e) => {
    });
  }
 
 public workbookSaveData(){
  const opt = new WorkbookSaveOptions();
    opt.type = "blob";
    this.spreadsheet.workbook.save(opt, (d) => {
      let fileReader = new FileReader();
      fileReader.readAsBinaryString(d as Blob);
      fileReader.onload = (e: any) => {
        debugger;
        var workbook = XLSX.read(fileReader.result, { type: 'binary' });
        var sheetNames = workbook.SheetNames;
        this.ExcelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
        this.workbookSaveInFolder();
      }
    }, (e) => {
    });
 }
  public workbookSaveInFolder(): void {
    this.appSerrvice.InsertDealerDetails("child2", JSON.stringify(this.ExcelData)).subscribe((response: any) => {
      const opt = new WorkbookSaveOptions();
      opt.type = "blob";
      this.spreadsheet.workbook.save(opt, (d) => {
        const formData = new FormData();
        formData.append('file', d as Blob, "Child2WorkbookData.xlsx");
        this.appSerrvice.getFileUpload(formData).subscribe(res => {
          alert("Inserted Records")
        });
      }, (e) => {
      });
    });
  }
  public getFileUpload(formData: FormData){

  }
  // public workbookSave(): void {
  //   const opt = new WorkbookSaveOptions();
  //   opt.type = "blob";
  //   this.spreadsheet.workbook.save(opt, (d) => {
  //     let fileReader = new FileReader();
  //     fileReader.readAsBinaryString(d as Blob);
  //     fileReader.onload = (e: any) => {
  //       var workbook = XLSX.read(fileReader.result, { type: 'binary' });
  //       var sheetNames = workbook.SheetNames;
  //       this.ExcelData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetNames[0]]);
  //       if(this.ExcelData != null){
  //         for (var i = 1; i < Object.keys(this.ExcelData[0]).length; i++) {
  //           var sumFormula = Formula.parse("=SUM(" + AlphaBetica[i] + "1:" + AlphaBetica[i] + "" + this.ExcelData.length + ")", CellReferenceMode.A1);
  //           sumFormula.applyTo(this.spreadsheet.activeWorksheet.rows(this.ExcelData.length).cells(i));
  //         }
  //       }
  //       this.appSerrvice.InsertDealerDetails("child2", JSON.stringify(this.ExcelData)).subscribe((response: any) => {
  //         this.workbookSaveInFolder();
  //       });
  //     }
  //   }, (e) => {
  //   });
  // }
  // public workbookSaveInFolder(): void {
  //   const opt = new WorkbookSaveOptions();
  //   opt.type = "blob";
  //   this.spreadsheet.workbook.save(opt, (d) => {
  //     const formData = new FormData();
  //     formData.append('file', d as Blob, "Child2WorkbookData.xlsx");
  //     this.appSerrvice.getFileUpload(formData).subscribe(res => {
  //       alert("Inserted Records")
  //     });
  //   }, (e) => {
  //   });
  // }
  public workbookDownload(): void {
    ExcelUtility.save(this.spreadsheet.workbook, ".xlsx");
  }
  public AddRowColumn() {
  }
  ngAfterViewInit(): void{
  //   this.spreadsheet.activeCellChanged.subscribe((f: any)=>{
  //     debugger;
  //     alert(f);
  //  })
  }
  
  public onChange() {

    this.spreadsheet.activeWorksheet.protect();
    this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = false;
    
    //this.spreadsheet.activeWorksheet.rows(3).Ins = false;
    // this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = false;
    // if(this.select.nativeElement.value == "101") {
    //   this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    // }
    // if(this.select.nativeElement.value == "102") {
    //   this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    // }
    // if(this.select.nativeElement.value == "103") {
    //   this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    // }
    // if(this.select.nativeElement.value == "104") {
    //   this.spreadsheet.activeWorksheet.rows(4).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(5).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(6).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(7).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(8).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(9).cellFormat.locked = true;
    //   this.spreadsheet.activeWorksheet.rows(10).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.rows(11).cellFormat.locked = false;
    //   this.spreadsheet.activeWorksheet.columns(6).cellFormat.locked = true;
    // }
    
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
