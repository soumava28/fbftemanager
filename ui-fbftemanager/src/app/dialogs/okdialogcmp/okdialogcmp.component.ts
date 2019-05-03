import { Component, OnInit, Inject } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialogRef, MatDialog } from '@angular/material';

@Component({
    selector: 'app-okdialogcmp',
    templateUrl: './okdialogcmp.component.html',
    styleUrls: ['./okdialogcmp.component.scss']
})
export class OkdialogcmpComponent {
     message = 'Are you sure?'; // Default messages
    confirmButtonText = 'Yes';
    cancelButtonText = 'Cancel';
    constructor(
        @Inject(MAT_DIALOG_DATA) private data: any,
        private dialogRef: MatDialogRef<OkdialogcmpComponent>) {
        if (data) {
            this.message = data.message || this.message;
            if (data.buttonText) {
                this.confirmButtonText = data.buttonText.ok || this.confirmButtonText;
                //this.cancelButtonText = data.buttonText.cancel || this.cancelButtonText;
            }
        }
        this.dialogRef.updateSize('300vw', '300vw');
    }

    onConfirmClick(): void {
        this.dialogRef.close(true);
    }
}
