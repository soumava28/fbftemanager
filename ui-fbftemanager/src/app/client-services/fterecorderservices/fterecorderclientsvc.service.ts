import { Injectable } from '@angular/core';
import { HttpClient, HttpHeaders, HttpParams } from '@angular/common/http';
import { FteRecordClientModel } from 'src/app/client-models/fteclientmodels/fterecordclientmodel';

@Injectable({
    providedIn: 'root'
})
export class FterecorderclientsvcService {

    postUrl = 'http://localhost:8090/ftemodifications/fterecorder/';
    getFteCountUrl = 'http://localhost:2311/fetchftedata/getTotalFTECountByTrack/';
    constructor(private client: HttpClient) { }

    recordIt(fteRec: FteRecordClientModel[]) {
        return this.client.post<any>(this.postUrl, fteRec);
    }

    getFteCount(track: string) {
        return this.client.get(this.getFteCountUrl + track);
    }
    updateFteData(record : FteRecordClientModel){
        console.log(record);
        return this.client.put(this.postUrl + 'updateRecord',  record).subscribe();
    }
}
