import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { DatePipe } from '@angular/common';

@Injectable({
    providedIn: 'root'
})
export class DefectfetchService {
    private fteDataRetrievedSet: any[] = [];
    displayedColumns: string[] = ['track'];
    constructor(private httpServices: HttpClient) { }
    baseUrl = 'http://localhost:8090/defectfetch/defectRetriever/';
    backUrl = '';
    getDataFromHttpSource(source): any {
        this.backUrl = 'fetchDefect';
        const params = new HttpParams().set('source', source);
        console.log(params);
        return this.httpServices.get(this.baseUrl + this.backUrl, { params });
    }

    getDataFromSource(source): any[] {
        this.getDataFromHttpSource(source).subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
            });
        return this.fteDataRetrievedSet;
    }


    getDataInArrayFormat(defectmodel): any[] {
        this.fteDataRetrievedSet = [];
        return this.getDataFromSource(defectmodel.source);
    }
}
