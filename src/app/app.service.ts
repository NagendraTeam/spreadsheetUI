import { HttpClient, HttpHeaders} from '@angular/common/http';
import { Injectable } from '@angular/core';

@Injectable({
  providedIn: 'root'
})
export class AppService {

  constructor(private httpClient: HttpClient) { }
  // InsertChild1Data(formdata: FormData){
  //   return this.httpClient.post('https://localhost:44328/api/workbook/InsertChild1Data',formdata);
  // }
  getParentData(data: string){
    debugger;
    let headers = new HttpHeaders ({ 'Content-Type': 'application/json' });
    return this.httpClient.post('https://localhost:44328/api/workbook/getParentFile?data=' + data , headers);
  }
  InsertDealerDetails(name: string, spreadsheetInfo: string){
    debugger;
    let headers = new HttpHeaders ({ 'Content-Type': 'application/json' });
    return this.httpClient.post('https://localhost:44328/api/workbook/InsertDealerDetails?name='+name+'&sheetInfo='+ spreadsheetInfo, headers);
  }
  InsertChild2Data(formdata: FormData){
    return this.httpClient.post('https://localhost:44328/api/workbook/InsertChild2Data',formdata);
  }
  getParentWorbookData() {
    debugger;
    return this.httpClient.get('https://localhost:44328/api/workbook/GetParentWorkbookData');
  }
  getChild1WorbookData() {
    return this.httpClient.get('https://localhost:44328/api/workbook/GetChild1WorkbookData');
  }
  getChild2WorbookData() {
    return this.httpClient.get('https://localhost:44328/api/workbook/GetChild2WorkbookData');
  }
  getFileUpload(formdata: FormData){
    return this.httpClient.post('https://localhost:44328/api/workbook/UploadChild1Data',formdata);
  }
  getFile() {
    return this.httpClient.get('https://localhost:44328/api/workbook/getFile',{responseType: 'blob'});
  }
  getChild1File() {
    return this.httpClient.get('https://localhost:44328/api/workbook/getChild1File',{responseType: 'blob'});
  }
  getChild2File() {
    return this.httpClient.get('https://localhost:44328/api/workbook/getChild2File',{responseType: 'blob'});
  }
  getParentFile() {
    return this.httpClient.get('https://localhost:44328/api/workbook/getParentFile',{responseType: 'blob'});
  }
  // getChild2WorbookData(formdata: FormData){
  //   this.httpClient.post('https://localhost:44328/api/workbook/InsertWorkbookData',formdata).
  //   toPromise().then(
  //     res => {
  //       console.log(res);
  //     }, 
  //     err => {
  //        console.log(err);
  //     }
  //   );
  // }
}
