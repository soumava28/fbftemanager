import { Component, OnInit, Inject } from '@angular/core';
import { FormsModule, ReactiveFormsModule } from '@angular/forms';
import { MAT_DIALOG_DATA, MatDialogRef, MatDialog } from '@angular/material';
@Component({
    selector: 'app-editdilalogcmp',
    templateUrl: './editdilalogcmp.component.html',
    styleUrls: ['./editdilalogcmp.component.scss']
})
export class EditdilalogcmpComponent {
    message = 'Are you sure?'; // Default messages
    confirmButtonText = 'Submit';
    cancelButtonText = 'Cancel';
    defectData: any;
    constructor(
        @Inject(MAT_DIALOG_DATA) private data: any,
        private dialogRef: MatDialogRef<EditdilalogcmpComponent>) {
        this.defectData = data.defect;
        this.dialogRef.updateSize('300vw', '300vw');
    }

    onConfirmClick(defect): void {

        this.dialogRef.close(true);
    }
}
