import { NgModule } from '@angular/core';
import { BrowserModule } from '@angular/platform-browser';
import { IgxExcelModule } from 'igniteui-angular-excel';
import { IgxSpreadsheetModule } from 'igniteui-angular-spreadsheet';
import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { AppService } from './app.service';
import { ExcelUtility } from './ExcelUtility';
import { HttpClientModule, HttpClient } from '@angular/common/http';
import { ParentComponent } from './Parent/parent/parent.component';
import { Child1Component } from './Child1/child1/child1.component';
import { Child2Component } from './Child2/child2/child2.component';  



@NgModule({
  declarations: [
    AppComponent,
    ParentComponent,
    Child1Component,
    Child2Component
  ],
  imports: [
    BrowserModule,
    AppRoutingModule,
    IgxExcelModule,
    IgxSpreadsheetModule,
    HttpClientModule  
  ],
  providers: [ExcelUtility,AppService],
  bootstrap: [AppComponent]
})
export class AppModule { }
