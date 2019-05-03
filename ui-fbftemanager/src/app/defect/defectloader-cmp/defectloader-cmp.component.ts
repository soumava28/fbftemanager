import { Component, OnInit } from '@angular/core';
import { HttpClient, HttpEventType, HttpResponse } from '@angular/common/http';

@Component({
    selector: 'app-defectloader-cmp',
    templateUrl: './defectloader-cmp.component.html',
    styleUrls: ['./defectloader-cmp.component.scss']
})
export class DefectloaderCmpComponent implements OnInit {
    fileData: File = null;
    percentDone: number;
    uploadSuccess: boolean;

    postUrl = 'http://localhost:9181/defectLoader/loadDefectDump/';

    uploadIt(fileInput: any) {
        this.client.post(this.postUrl, fileInput).subscribe(
            resp => {
                return resp;
            });
    }

    constructor(private client: HttpClient) { }

    ngOnInit() {
    }

    fileProgress(fileInput: any) {
        this.fileData = fileInput.target.files[0];
    }

    uploadAndProgress(files: File[]) {
        console.log(files)
        var formData = new FormData();
        Array.from(files).forEach(f => formData.append('file', f))


        this.client.post(this.postUrl, formData, { reportProgress: true, observe: 'events' })

            .subscribe(event => {
                if (event.type === HttpEventType.UploadProgress) {
                    this.percentDone = Math.round(100 * event.loaded / event.total);
                } else if (event instanceof HttpResponse) {
                    this.uploadSuccess = true;
                }
            });
    }

}
