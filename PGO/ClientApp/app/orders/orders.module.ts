import { NgModule } from '@angular/core';
import { WjGridModule } from 'wijmo/wijmo.angular2.grid';
import { WjGridSheetModule } from 'wijmo/wijmo.angular2.grid.sheet';
import { WjInputModule } from 'wijmo/wijmo.angular2.input';
import { CommonModule } from '@angular/common';
import { OrdersRoutingModule } from './orders-routing.module';
import { OrdersComponent } from './orders.component';
import { OrderListComponent } from './order-list/order-list.component';

@NgModule({
    imports: [WjInputModule, WjGridModule, WjGridSheetModule, 
    CommonModule,
    OrdersRoutingModule
  ],
  declarations: [OrdersComponent, OrderListComponent]
})
export class OrdersModule { }
