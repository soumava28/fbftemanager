import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import { DefectsearchclientModule } from 'src/app/client-models/defectmodel/defectsearchclient.module';
import { DefectfetchService } from 'src/app/client-services/defectfetch-services/defectfetch.service';
import { EditdilalogcmpComponent } from 'src/app/dialogs/editdilalogcmp/editdilalogcmp.component';
import { MatDialog } from '@angular/material';

@Component({
    selector: 'app-defectretriever-cmp',
    templateUrl: './defectretriever-cmp.component.html',
    styleUrls: ['./defectretriever-cmp.component.scss']
})
export class DefectretrieverCmpComponent implements OnInit {

    // tslint:disable-next-line:typedef-whitespace
    defectsearchModel: DefectsearchclientModule = new DefectsearchclientModule();
    defectData: any[] = [];
    constructor(private defectFetch: DefectfetchService, private okDialogue: MatDialog) { }

    ngOnInit() {
        this.defectsearchModel.source = 'MEDS';
    }

    editData(defect) {
        console.log(defect);
        this.okDialogue.open(EditdilalogcmpComponent, {
            data: {
                defect,
                buttonText: {
                    submit: 'Submit'
                }
            }
        });
    }
    searchData() {
        this.defectData = [];
        this.defectData = this.defectFetch.getDataInArrayFormat(this.defectsearchModel);
    }

    editDatatest() {


    }
}
