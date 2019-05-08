import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { DatePipe } from '@angular/common';
import { MatTableDataSource } from '@angular/material';

@Injectable({
    providedIn: 'root'
})
export class DefectfetchService {
    private fteDataRetrievedSet: any[] = [];
    displayedColumns: string[] = ['track'];
    constructor(private httpServices: HttpClient) { }
    baseUrl = 'http://localhost:8090/defectfetch/defectRetriever/';
    backUrl = '';
    exportUrl = 'http://localhost:8080/defectsToExcel/saveDataExcel';
    dataSource: MatTableDataSource<any> = new MatTableDataSource();
    getDataFromHttpSource(source): any {
        this.backUrl = 'fetchDefect';
        const params = new HttpParams().set('source', source);
        console.log(params);
        return this.httpServices.get(this.baseUrl + this.backUrl, { params });
    }

    getDataFromSource(source): MatTableDataSource<any> {
        this.getDataFromHttpSource(source).subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
                this.dataSource.data = this.fteDataRetrievedSet;
            });
        return this.dataSource;
    }


    getDataInArrayFormat(defectmodel): MatTableDataSource<any> {
        this.dataSource.data = [];
        return this.getDataFromSource(defectmodel.source);
    }

    exportDataToExcel(): string {
        return this.exportUrl;
    }
}
