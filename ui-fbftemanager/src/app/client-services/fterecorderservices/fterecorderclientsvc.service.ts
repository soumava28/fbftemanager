import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';
import { FteRecordClientModel } from 'src/app/client-models/fteclientmodels/fterecordclientmodel';

@Injectable({
    providedIn: 'root'
})
export class FterecorderclientsvcService {

    postUrl =  'http://localhost:9129/fterecorder/';
    constructor(private client: HttpClient) { }

    recordIt(fteRec: FteRecordClientModel[]) {
        return this.client.post<any>(this.postUrl, fteRec);
    }
}
