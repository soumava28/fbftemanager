import { BrowserModule } from '@angular/platform-browser';
import { NgModule } from '@angular/core';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { NavComponent } from './nav/nav.component';
import { FooterComponent } from './footer/footer.component';
import { HttpClientModule } from '@angular/common/http';
import { LoginCmpComponent } from './login/login-cmp/login-cmp.component';
import { DashboardModule } from './dashboard/dashboard.module';
import { DatePipe } from '@angular/common';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { EditdilalogcmpComponent } from './dialogs/editdilalogcmp/editdilalogcmp.component';
import { FormsModule } from '@angular/forms';
import { EditftedilougeComponent } from './dialogs/editftedilouge/editftedilouge.component';



@NgModule({
    declarations: [
        AppComponent,
        NavComponent,
        FooterComponent,
        LoginCmpComponent,
        EditdilalogcmpComponent,
        EditftedilougeComponent
    ],
    imports: [
        BrowserModule,
        AppRoutingModule,
        HttpClientModule,
        DashboardModule,
        FormsModule,
        BrowserAnimationsModule
    ],
    providers: [DatePipe],
    bootstrap: [AppComponent]
})
export class AppModule { }
