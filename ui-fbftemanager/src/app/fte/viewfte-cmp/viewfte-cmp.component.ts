import { Component, OnInit } from '@angular/core';
import { FteretrieverService } from 'src/app/client-services/fteretrieverservices/fteretriever.service';
import { MatDialog } from '@angular/material';
import { FteRecordClientModel } from 'src/app/client-models/fteclientmodels/fterecordclientmodel';

@Component({
    selector: 'app-viewfte-cmp',
    templateUrl: './viewfte-cmp.component.html',
    styleUrls: ['./viewfte-cmp.component.scss']
})
export class ViewfteCmpComponent implements OnInit {

    public fteDataAfterSearch: any[] = [];
    fterecordmodel: FteRecordClientModel = new FteRecordClientModel();
    constructor(private fteRetrieveServiece: FteretrieverService) {


    }
    searchData() {
        this.fteDataAfterSearch = [];
        console.log('Fetched Data --->');
        console.log(this.fterecordmodel);
        this.fteDataAfterSearch = this.fteRetrieveServiece.getDataInArrayFormat(this.fterecordmodel);
    }
    ngOnInit() {
    }

}
