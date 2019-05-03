import { NgModule } from '@angular/core';
import { CommonModule } from '@angular/common';
import { DashboardCmpComponent } from './dashboard-cmp/dashboard-cmp.component';
import { BrowserModule } from '@angular/platform-browser';
import { RouterModule, Routes } from '@angular/router';
import { FteCmpComponent } from '../fte/fte-cmp/fte-cmp.component';
import { ViewfteCmpComponent } from '../fte/viewfte-cmp/viewfte-cmp.component';
import { SidenavCmpComponent } from '../sidenav/sidenav-cmp/sidenav-cmp.component';
import { CustomMaterialModule } from '../material/material.module';
import { FteModule } from '../fte/fte.module';
import { FeedbackModule } from '../feedback/feedback.module';
import { ViewfeedbackCmpComponent } from '../feedback/viewfeedback-cmp/viewfeedback-cmp.component';
import { FeedbackCmpComponent } from '../feedback/feedback-cmp/feedback-cmp.component';
import { OkdialogcmpComponent } from '../dialogs/okdialogcmp/okdialogcmp.component';
import { DefectloaderCmpComponent } from '../defect/defectloader-cmp/defectloader-cmp.component';
import { DefectretrieverCmpComponent } from '../defect/defectretriever-cmp/defectretriever-cmp.component';
import { DefectModule } from '../defect/defect.module';
import { EditdilalogcmpComponent } from '../dialogs/editdilalogcmp/editdilalogcmp.component';

const appRoutes: Routes = [
    { path: 'dashboard/', component: FteCmpComponent, pathMatch: 'full' },
    {
        path: 'dashboard', component: DashboardCmpComponent,
        children: [
            { path: 'newfte', component: FteCmpComponent },
            { path: 'viewfte', component: ViewfteCmpComponent },
            { path: 'newfeedback', component: FeedbackCmpComponent },
            { path: 'viewfeedback', component: ViewfeedbackCmpComponent },
            { path: 'loaddefect', component: DefectloaderCmpComponent },
            { path: 'retrievedefect', component: DefectretrieverCmpComponent }
        ]
    }
];

@NgModule({
    declarations: [DashboardCmpComponent, SidenavCmpComponent, OkdialogcmpComponent  ],
    imports: [
        CommonModule,
        BrowserModule,
        RouterModule.forChild(
            appRoutes,
        ),
        FteModule,
        FeedbackModule,
        DefectModule,
        CustomMaterialModule,
    ],
    entryComponents: [OkdialogcmpComponent , EditdilalogcmpComponent],
})
export class DashboardModule { }
