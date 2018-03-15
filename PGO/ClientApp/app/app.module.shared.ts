import { NgModule } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { HttpModule } from '@angular/http';
import { RouterModule } from '@angular/router';
import { BrowserAnimationsModule } from '@angular/platform-browser/animations';
import { Headers, RequestOptions, BaseRequestOptions } from '@angular/http';
import { APP_BASE_HREF, CommonModule, Location, LocationStrategy, HashLocationStrategy } from '@angular/common';
// third party module to display toast 
import { ToastrModule } from 'toastr-ng2';
import { BrowserModule } from '@angular/platform-browser';
import { WjGridModule } from 'wijmo/wijmo.angular2.grid';
import { WjGridSheetModule } from 'wijmo/wijmo.angular2.grid.sheet';
import { WjInputModule } from 'wijmo/wijmo.angular2.input';
import { TreeModule, TreeNode } from 'angular-tree-component';
import { WjGridGrouppanelModule} from 'wijmo/wijmo.angular2.grid.grouppanel';
//PRIMENG - Third party module
import { InputTextModule, DataTableModule, ButtonModule, DialogModule } from 'primeng/primeng';

import { AppComponent } from './components/app/app.component';
import { NavMenuComponent } from './components/navmenu/navmenu.component';
import { FlexSheet } from './components/flexsheet/flexsheet.component';
import { ModalModule } from 'ngx-bootstrap/modal';

import { BudgetService } from './_services/index';

import { ContextmenuComponent } from './components/contextmenu/contextmenu.component';
import { ContextMenuModule } from 'angular2-contextmenu';

class AppBaseRequestOptions extends BaseRequestOptions {
    headers: Headers = new Headers();
    constructor() {
        super();
        this.headers.append('Content-Type', 'application/json');
        this.body = '';
    }
}

@NgModule({
    declarations: [
        AppComponent,
        NavMenuComponent,
        FlexSheet,
        ContextmenuComponent
    ],
    providers: [BudgetService,
        { provide: LocationStrategy, useClass: HashLocationStrategy },
        { provide: RequestOptions, useClass: AppBaseRequestOptions }],
    imports: [WjInputModule, WjGridModule, WjGridSheetModule, WjGridGrouppanelModule,BrowserModule,ContextMenuModule,
        CommonModule,
        HttpModule,
        FormsModule,
        BrowserAnimationsModule,
        ToastrModule.forRoot(),
        TreeModule,
        InputTextModule, DataTableModule, ButtonModule, DialogModule,
        ModalModule.forRoot(),
        RouterModule.forRoot([
            { path: '', redirectTo: 'sheet', pathMatch: 'full' },
            { path: 'sheet', component: FlexSheet },
            {
                path: 'orders', loadChildren:
                './orders/orders.module#OrdersModule'
            },
            { path: '**', redirectTo: 'sheet' },
           
        ])
    ],
    exports: [RouterModule]
})
export class AppModuleShared {
}
