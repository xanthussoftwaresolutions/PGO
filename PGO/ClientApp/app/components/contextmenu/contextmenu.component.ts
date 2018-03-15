import { Component, Input } from '@angular/core';

@Component({
    selector: 'app-contextmenu',
    templateUrl: './contextmenu.component.html',
    styleUrls: ['./contextmenu.component.css'],
})
export class ContextmenuComponent {

    constructor() { }


    @Input() x = 0;
    @Input() y = 0;

}