import { Injectable } from '@angular/core';
import { HttpClient } from '@angular/common/http';

@Injectable({
    providedIn: 'root'
})
export class DefectrecorderService {
    postUrl = 'http://localhost:8090/defectlist/defectLoader/updateDefectDump';
    constructor(private client: HttpClient) { }

    updateData(defect: any) {
        return this.client.put(this.postUrl, defect).subscribe();
    }
}
