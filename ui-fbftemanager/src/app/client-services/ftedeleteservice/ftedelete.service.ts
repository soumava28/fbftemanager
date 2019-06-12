import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { JsonPipe } from '@angular/common';

@Injectable({
    providedIn: 'root'
})
export class FtedeleteService {
    postUrl = 'http://localhost:8090/ftemodifications/fterecorder/';
    constructor(private client: HttpClient) { }

    removeData(id) {
        const params = new HttpParams().set('id', id);
        return this.client.get(this.postUrl + '/deleteRecord', { params });
    }
    updateFteData(record : any){
        console.log("new Data" + record);
        return this.client.put(this.postUrl + 'updateRecord',  record ).subscribe();
    }
}
