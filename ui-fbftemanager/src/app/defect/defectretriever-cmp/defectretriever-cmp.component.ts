import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import { DefectsearchclientModule } from 'src/app/client-models/defectmodel/defectsearchclient.module';
import { DefectfetchService } from 'src/app/client-services/defectfetch-services/defectfetch.service';
import { EditdilalogcmpComponent } from 'src/app/dialogs/editdilalogcmp/editdilalogcmp.component';
import { MatPaginator, MatSort, MatTableDataSource, MatDialog } from '@angular/material';

@Component({
    selector: 'app-defectretriever-cmp',
    templateUrl: './defectretriever-cmp.component.html',
    styleUrls: ['./defectretriever-cmp.component.scss']
})
export class DefectretrieverCmpComponent implements OnInit {

    // tslint:disable-next-line:typedef-whitespace
    defectsearchModel: DefectsearchclientModule = new DefectsearchclientModule();
    defectData: any[] = [];
    displayedColumns: string[] = ['track', 'Key', 'Summary',
        'applicableForIe?', 'tobefixed?', 'Comments', 'upDate', 'Edit'];
    dataSource = new MatTableDataSource();
    @ViewChild(MatPaginator) paginator: MatPaginator;
    @ViewChild(MatSort) sort: MatSort;
    exportUrl = 'http://localhost:8080/defectsToExcel/saveDataExcel';
    constructor(private defectFetch: DefectfetchService, private editDialog: MatDialog) { }

    ngOnInit() {
        this.defectsearchModel.source = 'Meds';
    }

    editData(defect) {
        console.log(defect);
      let isSaved = this.editDialog.open(EditdilalogcmpComponent, {
            data: {
                dataSource : this.dataSource,
                defect,
                buttonText: {
                    submit: 'Submit'
                }
            }
        });
        isSaved.afterClosed().subscribe(result => {
            if (result) {
                this.searchData();
            }
        });
    }
    searchData() {
        this.dataSource = new MatTableDataSource();
        this.dataSource = this.defectFetch.getDataInArrayFormat(this.defectsearchModel);
        this.dataSource.sort = this.sort;
        this.dataSource.paginator = this.paginator;
    }


    applyFilter(filterValue: string) {
        this.dataSource.filter = filterValue.trim().toLowerCase();
    }

    getUrltoExport(): string {
        return this.defectFetch.exportDataToExcel();
    }
}
