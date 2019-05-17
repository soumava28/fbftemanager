import { Component, OnInit, ViewChild, ElementRef } from '@angular/core';
import { FteretrieverService } from 'src/app/client-services/fteretrieverservices/fteretriever.service';
import { FteRecordClientModel } from 'src/app/client-models/fteclientmodels/fterecordclientmodel';
import * as XLSX from 'xlsx';
import { MatTableDataSource, MatDialog, MatPaginator, MatSort } from '@angular/material';
import { FtedeleteService } from 'src/app/client-services/ftedeleteservice/ftedelete.service';
import { OkdialogcmpComponent } from 'src/app/dialogs/okdialogcmp/okdialogcmp.component';


@Component({
    selector: 'app-viewfte-cmp',
    templateUrl: './viewfte-cmp.component.html',
    styleUrls: ['./viewfte-cmp.component.scss']
})

export class ViewfteCmpComponent implements OnInit {

    public fteDataAfterSearch: any[];
    fterecordmodel: FteRecordClientModel = new FteRecordClientModel();

    @ViewChild('TABLE') table: ElementRef;
    @ViewChild(MatPaginator) paginator: MatPaginator;
    @ViewChild(MatSort) sort: MatSort;
    public displayedColumns: string[] = ['track', 'weekStDt', 'weekEdDt', 'ftesPerRosterCount', 'defectCount', 'workableDefectsCount',
        'defectFteCount', 'widgetCount', 'ftesLoanedCount', 'ptoCount', 'ftesBorrowedCount', 'ftesForPerformanceCount',
        'ftesForExtendedScenarioExecCount', 'medsDefAnalysisCount', 'excessFteCount', 'note', 'delete'];
    public dataSource = new MatTableDataSource();


    ExportTOExcel() {
        const ws: XLSX.WorkSheet = XLSX.utils.table_to_sheet(this.table.nativeElement);
        const wb: XLSX.WorkBook = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

        /* save to file */
        XLSX.writeFile(wb, 'SheetJS.xlsx');

    }


    // tslint:disable-next-line:max-line-length
    constructor(private fteRetrieveServiece: FteretrieverService, private okDialogue: MatDialog, private ftedeleteService: FtedeleteService) { }

    searchData() {
        this.dataSource = null;
        this.dataSource = this.fteRetrieveServiece.getDataInArrayFormat(this.fterecordmodel);
        this.dataSource.sort = this.sort;
        this.dataSource.paginator = this.paginator;
    }
    ngOnInit() {
        this.fterecordmodel.weekEdDt = new Date();
        this.fterecordmodel.weekStDt = new Date();
        this.fterecordmodel.track = 'All';
    }
    deleteData(id) {
        this.ftedeleteService.removeData(id)
            .subscribe(
                data => {
                    console.log('Success!', data);
              let deleteDone =  this.okDialogue.open(OkdialogcmpComponent, {
                        data: {
                            message: 'Deleted Successfully!',
                            buttonText: {
                                ok: 'OK'
                                // cancel: 'No'
                            }
                        }

                    });
                    deleteDone.afterClosed().subscribe(result => {
                        if (result) {
                            this.searchData();
                        }
                    });
                },
                error => {
                    console.error('Error!', error);
                    this.okDialogue.open(OkdialogcmpComponent, {
                        data: {
                            message: 'Something went wrong! Please try again/Contact the administrator!',
                            buttonText: {
                                ok: 'OK'
                                // cancel: 'No'
                            }
                        }
                    });
                }
            );
        
    }
}

