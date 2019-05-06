import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DefectloaderCmpComponent } from './defectloader-cmp/defectloader-cmp.component';
import { DefectretrieverCmpComponent } from './defectretriever-cmp/defectretriever-cmp.component';
import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { MatTableModule, MatPaginator, MatPaginatorModule, MatFormFieldModule, MatInputModule } from '@angular/material';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
@NgModule({
    declarations: [DefectloaderCmpComponent, DefectretrieverCmpComponent],
    imports: [
        CommonModule,
        FormsModule,
        BrowserModule,
        BrowserAnimationsModule,
        MatTableModule,
        ReactiveFormsModule,
        MatPaginatorModule,
        MatFormFieldModule,
        MatInputModule
    ]
})
export class DefectModule {

}
