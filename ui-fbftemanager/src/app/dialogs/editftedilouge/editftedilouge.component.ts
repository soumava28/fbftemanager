import { Component, OnInit, Inject } from '@angular/core';
import { MAT_DIALOG_DATA, MatDialogRef, MatDialog } from '@angular/material';
import { DefectrecorderService } from 'src/app/client-services/defectrecorder/defectrecorder.service';
import { FterecorderclientsvcService } from 'src/app/client-services/fterecorderservices/fterecorderclientsvc.service';
import { FtedeleteService } from 'src/app/client-services/ftedeleteservice/ftedelete.service';
@Component({
  selector: 'app-editftedilouge',
  templateUrl: './editftedilouge.component.html',
  styleUrls: ['./editftedilouge.component.scss']
})
export class EditftedilougeComponent implements OnInit {

  ngOnInit() {
  }
  message = 'Are you sure?'; // Default messages
  confirmButtonText = 'Submit';
  cancelButtonText = 'Cancel';
  editData: any;
  isEdited :any;
  constructor(
      @Inject(MAT_DIALOG_DATA) private data: any,
      private dialogRef: MatDialogRef<EditftedilougeComponent>, private fteUpdate: FtedeleteService) {
      this.editData = data.ftedate;
      this.dialogRef.updateSize('300vw', '300vw');
  }

  onConfirmClick(editData) {
      this.fteUpdate.updateFteData(editData);
      this.dialogRef.close(true);
  }

}
