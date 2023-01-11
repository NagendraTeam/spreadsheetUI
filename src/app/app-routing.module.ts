import { NgModule } from '@angular/core';
import { RouterModule, Routes } from '@angular/router';
import { Child1Component } from './Child1/child1/child1.component';
import { Child2Component } from './Child2/child2/child2.component';
import { ParentComponent } from './Parent/parent/parent.component';

const routes: Routes = [
  { path: '', component: ParentComponent },
  { path: 'child1', component: Child1Component },
  { path: 'child2', component: Child2Component }
];

@NgModule({
  imports: [RouterModule.forRoot(routes)],
  exports: [RouterModule]
})
export class AppRoutingModule { }
