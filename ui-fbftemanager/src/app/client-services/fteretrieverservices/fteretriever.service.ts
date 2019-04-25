import { Injectable } from '@angular/core';
import { HttpClient, HttpParams } from '@angular/common/http';
import { DatePipe } from '@angular/common';
@Injectable({
    providedIn: 'root'
})
export class FteretrieverService {

    constructor(private httpServices: HttpClient, private datePipe: DatePipe) { }
    private fteDataRetrievedSet: any[] = [];
    private backUrl = '';

    baseUrl = 'http://localhost:8090/ftefetch/fetchftedata/';
    getDataAfterSearch(): any {
        return this.httpServices.get(this.baseUrl);
    }

    getDatafromWeekStDt(weekStDt): any {
        this.backUrl = 'byDate/';
        const params = new HttpParams().set('weekStDt', weekStDt);
        console.log(params);
        return this.httpServices.get(this.baseUrl + this.backUrl, { params });
    }


    getDatafromTrack(track): any {
        this.backUrl = 'byTrack/';
        const params = new HttpParams().set('track', track);
        console.log(params);
        return this.httpServices.get(this.baseUrl + this.backUrl, { params });
    }


    getDataWeekStDtnTrack(track, weekStDt): any {
        this.backUrl = 'byDatenTrack/';
        const params = new HttpParams().set('weekStDt', weekStDt).set('track', track);
        console.log(params);
        return this.httpServices.get(this.baseUrl + this.backUrl, { params });
    }

    getDatafromTrackArray(track): any[] {
        this.getDatafromTrack(track).subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
            });
        return this.fteDataRetrievedSet;
    }
    getDataFromWeekStDtArray(weekStDt): any[] {
        this.getDatafromWeekStDt(this.datePipe.transform(new Date(weekStDt), 'yyyy-MM-dd')).subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
            });
        return this.fteDataRetrievedSet;
    }

    getDataWeekStDtnTrackArray(track, weekStDt): any[] {
        this.getDataWeekStDtnTrack(track, this.datePipe.transform(new Date(weekStDt), 'yyyy-MM-dd')).subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
            });
        return this.fteDataRetrievedSet;
    }

    getAllDataInArrayFormat(): any[] {
        this.getDataAfterSearch().subscribe(
            resp => {
                Object.assign(this.fteDataRetrievedSet, resp);
            });
        return this.fteDataRetrievedSet;
    }

    getDataInArrayFormat(fterecordmodel): any[] {
        this.backUrl = '';
        if (fterecordmodel.track === undefined && fterecordmodel.weekStDt === undefined) {
            return this.getAllDataInArrayFormat();
        } else {
            if (fterecordmodel.track === undefined) {
                return this.getDataFromWeekStDtArray(fterecordmodel.weekStDt);
            } else if (fterecordmodel.weekStDt === undefined) {
                return this.getDatafromTrackArray(fterecordmodel.track);
            } else {
                return this.getDataWeekStDtnTrackArray(fterecordmodel.track, fterecordmodel.weekStDt);
            }
        }

    }
}
