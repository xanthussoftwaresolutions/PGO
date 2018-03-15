import * as wjcGridSheet from 'wijmo/wijmo.grid.sheet';
import * as wjcInput from 'wijmo/wijmo.input';
import * as wjcGrid from 'wijmo/wijmo.grid';
import * as wjcCore from 'wijmo/wijmo';





// Angular

import { Component, EventEmitter, Inject, enableProdMode, ViewChild, NgModule, OnInit, TemplateRef } from '@angular/core';
import { FormsModule } from '@angular/forms';
import { CommonModule } from '@angular/common';
import { platformBrowserDynamic } from '@angular/platform-browser-dynamic';
import { BudgetService } from '../../_services/index';
import { TREE_ACTIONS, KEYS, IActionMapping, ITreeOptions, TreeNode, TreeModel } from 'angular-tree-component';
import { Router } from '@angular/router';
import { Year } from '../../_models/index';
import { ContextMenuService, ContextMenuComponent } from 'angular2-contextmenu';
import { BsModalService } from 'ngx-bootstrap/modal';
import { BsModalRef } from 'ngx-bootstrap/modal/bs-modal-ref.service';
'use strict';
class YearInfo implements Year {
    constructor(public category?, public Class?, public year?, public id?, public jan?, public feb?, public march?, public april?, public may?, public june?, public july?, public aug?, public sept?, public oct?, public nov?, public dec?, public total?) { }
}
// The Explorer application root component.
@Component({
    selector: 'flexsheet',
    templateUrl: './flexsheet.component.html',
    styleUrls: ['./flexsheet.component.css'],
})
export class FlexSheet {
    expandType: any;
    public edited = false;
    public isGrowthRate = false;
    public isPercent = false;
    CategoryName: any;
    nodeItem: any;
    treeItem: any;
    EntryType = 'ManualEntry';
    EntryLastYear = 'LastYearSamePeriod';
    growthRate = { name: 'Jan 2018', value: 91 };
    growthRateLYSP = { name: '2018', value: 106 };
    StartLastYearSP = [{ name: '2018', value: 106 }, { name: '2019', value: 121 }, { name: '2020', value: 136 }, { name: '2021', value: 139 }, { name: '2022', value: 142 }, { name: '2023', value: 145 }, { name: '2024', value: 148 }, { name: '2025', value: 151 }, { name: '2026', value: 154 }];
    StartGrowthOver = [];
    args: any;
    percentArgs: any;
    isGrowthconfirm: any;
    ispercentconfirm: any;
    EntryPrior = "Current";
    //data: any[];
    data: wjcCore.CollectionView;
    sortManager: wjcGridSheet.SortManager;
    columns: string[];
    fonts: any[];
    fontSizeList: any[];
    selectionFormatState: wjcGridSheet.IFormatState;
    selection: any = {
        content: '',
        position: '',
        fontFamily: 'Arial, Helvetica, sans-serif',
        fontSize: '8px'
    };
    mergeState: any;
    isFrozen: boolean = false;
    undoStack: wjcGridSheet.UndoStack;
    currentCellData: any;
    fileName: string;
    tableStyleNames = null;
    selectedTable: wjcGridSheet.Table = null;

    private _format = '';
    private _updatingSelection = false;
    private _applyFillColor = false;
    deleteCat: Number[] = [];
    validCat: boolean;

    // references FlexSheet named 'formulaSheet' in the view
    @ViewChild('flexSheet') flexSheet: wjcGridSheet.FlexSheet;
    @ViewChild('formulaSheet') formulaSheet: wjcGridSheet.FlexSheet;
    @ViewChild('formulaSheet1') formulaSheet1: wjcGridSheet.FlexSheet;
    @ViewChild('growthRateSheet') growthRateSheet: wjcGridSheet.FlexSheet;
    @ViewChild('percentRateSheet') percentRateSheet: wjcGridSheet.FlexSheet;
    @ViewChild('validatemessage') validatemessage;
    @ViewChild('percentCheckPopup') percentCheckPopup;
    @ViewChild('catDiv') catDiv;
    @ViewChild('catDivInner') catDivInner;


    lastIdRoot1: number = 1;
    lastIdRoot2: number = 100;
    lastRowRoot1: number = 10;
    lastRowRoot2: number = 16;
    RowCount: number = 2000;
    MenuRowCount: number = 18;
    yearData: Array<Year> = new Array<YearInfo>();
    yearDataArr: Array<Array<Year>> = new Array<Array<YearInfo>>();
    year: Year = new YearInfo();

    nodes11: any = [];
    nodesPerCat: any = [];
    contextmenu = false;
    contextmenuX = 0;
    contextmenuY = 0;
    public items = [
        { name: 'John', otherProperty: 'Foo' },
        { name: 'Joe', otherProperty: 'Bar' }
    ];
    @ViewChild(ContextMenuComponent) public basicMenu: ContextMenuComponent;
    ngOnInit() {

        this.budgetService.getExcelCount().subscribe(response => {
            this.MenuRowCount = this.RowCount = parseInt(response['result']['item1']) + 4;
            //this.MenuRowCount = parseInt(response['result']['item1']) + 4;
            this.lastRowRoot2 = this.RowCount;
            this.nodes11 = response['result']['item2'];
            this.nodesPerCat = response['result']['item2'];
            this.lastRowRoot1 = this.lastRowRoot2 - this.nodes11[1].children.length - 1;
        })
        let year = 2018, months = 'Jan';
        for (var i = 91; i < 135; i++) {
            this.StartGrowthOver.push({ name: months + ' ' + year, value: i });
            months = months == 'Jan' ? 'Feb' : months == 'Feb' ? 'March' : months == 'March' ? 'April' : months == 'April' ? 'May' : months == 'May' ? 'June' : months == 'June' ? 'July' : months == 'July' ? 'Aug' : months == 'Aug' ? 'Sept' : months == 'Sept' ? 'Oct' : months == 'Oct' ? 'Nov' : months == 'Nov' ? 'Dec' : 'Jan';
            year = months == 'Jan' ? year + 1 : year;
            i = months == 'Jan' ? i + 3 : i;
        }
        this.growthRate = this.StartGrowthOver[0];
        this.growthRateLYSP = this.StartLastYearSP[0];
    };
    treeInit(tree: any) {
        setTimeout(() => { tree.treeModel.expandAll(); }, 1000)
    }

    onEvent(event: any) {
        console.log(event);
    }
    options: ITreeOptions = {
        actionMapping: {
            mouse: {
                expanderClick: (tree, node, $event) => {
                    this.expandtreeDetails(tree, node, $event);

                },

            },


        }
    }
    addNode(newCat: any, messagetemp: TemplateRef<any>, $event) {
        let scrollposition = new wjcCore.Point(this.flexSheet.scrollPosition.x);
        if (newCat.trim() == "" || newCat == undefined) {
            this.modalRef = this.modalService.show(messagetemp, { class: 'modal-xs' });
            return;
        }
        this.ISvalidCat(newCat);
        if (!this.validCat) { return; }
        if (this.EntryType == 'ManualEntry') {
            if (this.nodeItem.data.id == 1) {
                this.flexSheet.insertRows(this.lastRowRoot1, 1);
                this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                for (var i = 1; i < 157; i++) {
                    if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                        this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                    }
                    else if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                        this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                    }
                    else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                        this.flexSheet.setCellData(this.lastRowRoot1, i, 0);
                    }
                }

                this.nodes11[0].children.push({
                    id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                    name: newCat,
                    ischecked: false
                });
                this.flexSheet.applyCellsStyle({
                    backgroundColor: '#ffff99'
                }, [new wjcGrid.CellRange(this.lastRowRoot1, 0, this.lastRowRoot1, 12), new wjcGrid.CellRange(this.lastRowRoot1, 15, this.lastRowRoot1, 27),
                new wjcGrid.CellRange(this.lastRowRoot1, 30, this.lastRowRoot1, 42),
                new wjcGrid.CellRange(this.lastRowRoot1, 45, this.lastRowRoot1, 57),
                new wjcGrid.CellRange(this.lastRowRoot1, 60, this.lastRowRoot1, 72),
                new wjcGrid.CellRange(this.lastRowRoot1, 75, this.lastRowRoot1, 87),
                new wjcGrid.CellRange(this.lastRowRoot1, 90, this.lastRowRoot1, 102),
                new wjcGrid.CellRange(this.lastRowRoot1, 105, this.lastRowRoot1, 117),
                new wjcGrid.CellRange(this.lastRowRoot1, 120, this.lastRowRoot1, 132),
                new wjcGrid.CellRange(this.lastRowRoot1, 135, this.lastRowRoot1, 157)]);
                this.treeItem.treeModel.update();
                this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
            }
            else if (this.nodeItem.data.id == 100) {
                this.flexSheet.insertRows(this.lastRowRoot2, 1);
                //this.flexSheet.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                for (var i = 1; i < 157; i++) {
                    if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                        this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                    }
                    else if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                        this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                    }
                    else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                        this.flexSheet.setCellData(this.lastRowRoot2, i, 0);
                    }
                }

                this.nodes11[1].children.push({
                    id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                    name: newCat,
                    ischecked: false
                });
                this.flexSheet.applyCellsStyle({
                    backgroundColor: '#ffff99'
                }, [new wjcGrid.CellRange(this.lastRowRoot2, 0, this.lastRowRoot2, 12), new wjcGrid.CellRange(this.lastRowRoot2, 15, this.lastRowRoot2, 27),
                new wjcGrid.CellRange(this.lastRowRoot2, 30, this.lastRowRoot2, 42),
                new wjcGrid.CellRange(this.lastRowRoot2, 45, this.lastRowRoot2, 57),
                new wjcGrid.CellRange(this.lastRowRoot2, 60, this.lastRowRoot2, 72),
                new wjcGrid.CellRange(this.lastRowRoot2, 75, this.lastRowRoot2, 87),
                new wjcGrid.CellRange(this.lastRowRoot2, 90, this.lastRowRoot2, 102),
                new wjcGrid.CellRange(this.lastRowRoot2, 105, this.lastRowRoot2, 117),
                new wjcGrid.CellRange(this.lastRowRoot2, 120, this.lastRowRoot2, 132),
                new wjcGrid.CellRange(this.lastRowRoot2, 135, this.lastRowRoot2, 157)]);
                this.treeItem.treeModel.update();
                this.lastRowRoot2 += 1;
            }
            this.edited = false;
            this.CategoryName = '';
            this._applyStyleForExpenceReportCss(this.formulaSheet1);
        }
        else if (this.EntryType == 'Percentage') {
            let validpercentCheck = false;

            this.nodesPerCat.forEach(obj => {
                obj.children.forEach(objInner => {
                    if (objInner.ischecked == true) {
                        validpercentCheck = true;
                    }
                });
            });
            if (!validpercentCheck) {
                this.modalRef = this.modalService.show(this.percentCheckPopup);
                return false;
            }
            if (this.EntryPrior == "Current") {
                if (this.nodeItem.data.id == 1) {
                    this.flexSheet.insertRows(this.lastRowRoot1, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    for (var i = 1, j = 2, k = 'A', l = 65, m = 1; i < 157; i++) {
                        if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                        }
                        else
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                            }
                            else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                    l += 1;
                                    m = m - 26;
                                }
                                else if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                    l += 1;
                                    m = 0;
                                }
                                if (k.length == 1 && k == 'Z') {
                                    m = 0;
                                }
                                if (k.length == 1 && k != 'Z') {
                                    k = String.fromCharCode(65 + m);
                                }
                                else {
                                    k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                }
                                let formula = '';
                                this.nodesPerCat.forEach(obj => {
                                    obj.children.forEach(objInner => {
                                        if (obj.id == 1 && objInner.ischecked == true) {
                                            formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                        } else if (obj.id == 100 && objInner.ischecked == true) {
                                            formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                        }
                                    });
                                });
                                if (formula != '') {
                                    formula += ',0)';
                                }
                                if (i >= 91) {
                                    this.flexSheet.setCellData(this.lastRowRoot1, i, formula.replace("+", "=round("));
                                    j++;
                                }

                            }
                        m++;
                    }
                    this.nodes11[0].children.push({
                        id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    this.treeItem.treeModel.update();
                    this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
                }
                else if (this.nodeItem.data.id == 100) {
                    this.flexSheet.insertRows(this.lastRowRoot2, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    for (var i = 1, j = 2, k = 'A', l = 65, m = 1; i < 157; i++) {
                        if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                        }
                        else
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                            }
                            else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                    l += 1;
                                    m = m - 26;
                                }
                                else if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                    l += 1;
                                    m = 0;
                                }
                                if (k.length == 1 && k == 'Z') {
                                    m = 0;
                                }
                                if (k.length == 1 && k != 'Z') {
                                    k = String.fromCharCode(65 + m);
                                }
                                else {
                                    k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                }
                                let formula = '';
                                this.nodesPerCat.forEach(obj => {
                                    obj.children.forEach(objInner => {
                                        if (obj.id == 1 && objInner.ischecked == true) {
                                            formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                        } else if (obj.id == 100 && objInner.ischecked == true) {
                                            formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                        }
                                    });
                                });
                                if (formula != '') {
                                    formula += ',0)';
                                }
                                if (i >= 91) {
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, formula.replace("+", "=round("));
                                    j++;
                                }
                            }
                        m++;
                    }
                    this.nodes11[1].children.push({
                        id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    this.treeItem.treeModel.update();
                    this.lastRowRoot2 += 1;
                }
                this.edited = false;
                this.CategoryName = '';
                this._applyStyleForExpenceReportCss(this.formulaSheet1);
            }
            else if (this.EntryPrior == "Prior") {
                if (this.nodeItem.data.id == 1) {
                    this.flexSheet.insertRows(this.lastRowRoot1, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    var minVal = 2, k = 'B', l = 65, m = 2
                    minVal = 2;
                    if (minVal <= 26) {
                        m = minVal, l = 65;
                    } else {
                        m = (minVal % 26), l = 64 + Math.floor(minVal / 26);
                    }
                    if (Math.floor(minVal / 26) == 0) {
                        k = String.fromCharCode(64 + m);
                    }
                    else {
                        k = String.fromCharCode(l) + String.fromCharCode(64 + m);
                    }
                    for (var i = 2, j = 2; i < 157; i++) {
                        if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                        }
                        else if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                        }
                        else if (i >= minVal) {
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                            }
                            else if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                            }
                            else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                let formula = '';

                                if (i >= 91) {
                                    if (k == 'EC') {
                                        this.nodesPerCat.forEach(obj => {
                                            obj.children.forEach(objInner => {
                                                if (obj.id == 1 && objInner.ischecked == true) {
                                                    formula += '+(SUM(DR' + (obj.children.indexOf(objInner) + 6) + ':EC' + (obj.children.indexOf(objInner) + 6) + ')* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                } else if (obj.id == 100 && objInner.ischecked == true) {
                                                    formula += '+(SUM(DR' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + ':EC' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + ')* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                }
                                            });
                                        });
                                        if (formula != '') {
                                            formula += ',0)';
                                        }
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, formula.replace("+", "=round("));
                                        j++;
                                    }
                                    else {
                                        this.nodesPerCat.forEach(obj => {
                                            obj.children.forEach(objInner => {
                                                if (obj.id == 1 && objInner.ischecked == true) {
                                                    formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                } else if (obj.id == 100 && objInner.ischecked == true) {
                                                    formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                }
                                            });
                                        });
                                        if (formula != '') {
                                            formula += ',0)';
                                        }
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, formula.replace("+", "=round("));
                                        j++;
                                    }
                                }
                                if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                    l += 1;
                                    m = m - 26;
                                }
                                else if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                    l += 1;
                                    m = 0;
                                }
                                if (k.length == 1 && k == 'Z') {
                                    m = 0;
                                }
                                if (k.length == 1 && k != 'Z') {
                                    k = String.fromCharCode(65 + m);
                                }
                                else {
                                    k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                }
                            }
                            m++;
                        }
                    }
                    this.nodes11[0].children.push({
                        id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    this.treeItem.treeModel.update();
                    this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
                }
                else if (this.nodeItem.data.id == 100) {
                    this.flexSheet.insertRows(this.lastRowRoot2, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    var minVal = 2, k = 'B', l = 65, m = 2
                    minVal = 1 + 1;
                    if (minVal <= 26) {
                        m = minVal, l = 65;
                    } else {
                        m = (minVal % 26), l = 64 + Math.floor(minVal / 26);
                    }
                    if (Math.floor(minVal / 26) == 0) {
                        k = String.fromCharCode(64 + m);
                    }
                    else {
                        k = String.fromCharCode(l) + String.fromCharCode(64 + m);
                    }
                    for (var i = 2, j = 2; i < 157; i++) {
                        if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                        } else if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                        }
                        else if (i >= minVal) {
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                            } else

                                if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                                }
                                else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                    let formula = '';

                                    if (i >= 91) {
                                        if (k == 'EC') {
                                            this.nodesPerCat.forEach(obj => {
                                                obj.children.forEach(objInner => {
                                                    if (obj.id == 1 && objInner.ischecked == true) {
                                                        formula += '+(SUM(DR' + (obj.children.indexOf(objInner) + 6) + ':EC' + (obj.children.indexOf(objInner) + 6) + ')* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                    } else if (obj.id == 100 && objInner.ischecked == true) {
                                                        formula += '+(SUM(DR' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + ':EC' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + ')* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                    }
                                                });
                                            });
                                            if (formula != '') {
                                                formula += ',0)';
                                            }
                                            this.flexSheet.setCellData(this.lastRowRoot2, i, formula.replace("+", "=round("));
                                            j++;
                                        }
                                        else {
                                            this.nodesPerCat.forEach(obj => {
                                                obj.children.forEach(objInner => {
                                                    if (obj.id == 1 && objInner.ischecked == true) {
                                                        formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                    } else if (obj.id == 100 && objInner.ischecked == true) {
                                                        formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                                    }
                                                });
                                            });
                                            if (formula != '') {
                                                formula += ',0)';
                                            }
                                            this.flexSheet.setCellData(this.lastRowRoot2, i, formula.replace("+", "=round("));
                                            j++;
                                        }
                                    }
                                    if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                        l += 1;
                                        m = m - 26;
                                    }
                                    else if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                        l += 1;
                                        m = 0;
                                    }
                                    if (k.length == 1 && k == 'Z') {
                                        m = 0;
                                    }
                                    if (k.length == 1 && k != 'Z') {
                                        k = String.fromCharCode(65 + m);
                                    }
                                    else {
                                        k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                    }
                                }
                            m++;
                        }
                    }
                    this.nodes11[1].children.push({
                        id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    this.treeItem.treeModel.update();
                    this.lastRowRoot2 += 1;

                }

                this.edited = false;
                this.CategoryName = '';
                this._applyStyleForExpenceReportCss(this.formulaSheet1);
            }
            else if (this.EntryPrior == "SubseQuent") {
                if (this.nodeItem.data.id == 1) {
                    this.flexSheet.insertRows(this.lastRowRoot1, 1);
                    this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    var k = 'EY';
                    for (var i = 156, j = 44; i > 1; i--) {
                        if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                        }
                        else if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                        }
                        else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                            let formula = '';

                            if (i >= 91 && i <= 154) {
                                if (k == 'EC') {
                                    this.nodesPerCat.forEach(obj => {
                                        obj.children.forEach(objInner => {
                                            if (obj.id == 1 && objInner.ischecked == true) {
                                                formula += '+((' + k + '' + (obj.children.indexOf(objInner) + 6) + '/12)* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            } else if (obj.id == 100 && objInner.ischecked == true) {
                                                formula += '+((' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + '/12)* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            }
                                        });
                                    });
                                    if (formula != '') {
                                        formula += ',0)';
                                    }
                                    this.flexSheet.setCellData(this.lastRowRoot1, i, formula.replace("+", "=round("));
                                    j--;
                                }
                                else {
                                    this.nodesPerCat.forEach(obj => {
                                        obj.children.forEach(objInner => {
                                            if (obj.id == 1 && objInner.ischecked == true) {
                                                formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            } else if (obj.id == 100 && objInner.ischecked == true) {
                                                formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 3) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            }
                                        });
                                    });
                                    if (formula != '') {
                                        formula += ',0)';
                                    }
                                    this.flexSheet.setCellData(this.lastRowRoot1, i, formula.replace("+", "=round("));
                                    j--;
                                }
                                k = wjcGridSheet.FlexSheet.convertNumberToAlpha(i);
                            }
                        }
                    }
                    this.nodes11[0].children.push({
                        id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    //  Applying Color for empty cell
                    //this.flexSheet.applyCellsStyle({
                    //    backgroundColor: '#ffff99'
                    //}, [new wjcGrid.CellRange(this.lastRowRoot1, 0, this.lastRowRoot1, 84), new wjcGrid.CellRange(this.lastRowRoot1, 145, this.lastRowRoot1, 145)]);
                    this.treeItem.treeModel.update();
                    this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
                }
                else if (this.nodeItem.data.id == 100) {
                    this.flexSheet.insertRows(this.lastRowRoot2, 1);
                    this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    var k = 'EY';
                    for (var i = 156, j = 44; i > 1; i--) {
                        if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                        }
                        else if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                        }
                        else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                            let formula = '';

                            if (i >= 91 && i <= 154) {
                                if (k == 'EC') {
                                    this.nodesPerCat.forEach(obj => {
                                        obj.children.forEach(objInner => {
                                            if (obj.id == 1 && objInner.ischecked == true) {
                                                formula += '+((' + k + '' + (obj.children.indexOf(objInner) + 6) + '/12)* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            } else if (obj.id == 100 && objInner.ischecked == true) {
                                                formula += '+((' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + '/12)* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            }
                                        });
                                    });
                                    if (formula != '') {
                                        formula += ',0)';
                                    }
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, formula.replace("+", "=round("));
                                    j--;
                                }
                                else {
                                    this.nodesPerCat.forEach(obj => {
                                        obj.children.forEach(objInner => {
                                            if (obj.id == 1 && objInner.ischecked == true) {
                                                formula += '+(' + k + '' + (obj.children.indexOf(objInner) + 6) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            } else if (obj.id == 100 && objInner.ischecked == true) {
                                                formula += '+(' + k + '' + (obj.children.indexOf(objInner) + this.lastRowRoot1 + 2) + '* ' + this.percentRateSheet.getCellValue(j, 1) + ' / 100)';
                                            }
                                        });
                                    });
                                    if (formula != '') {
                                        formula += ',0)';
                                    }
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, formula.replace("+", "=round("));
                                    j--;
                                }
                                k = wjcGridSheet.FlexSheet.convertNumberToAlpha(i);
                            }

                        }
                    }
                    this.nodes11[1].children.push({
                        id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });

                    this.treeItem.treeModel.update();
                    this.lastRowRoot2 += 1;

                }

                this.edited = false;
                this.CategoryName = '';
                this._applyStyleForExpenceReportCss(this.formulaSheet1);
            }
        }
        else if (this.EntryType == 'GrowthRate') {
            if (this.EntryLastYear == 'LastYearSamePeriod') {
                if (this.nodeItem.data.id == 1) {
                    this.flexSheet.insertRows(this.lastRowRoot1, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    for (var i = 1, j = 2, k = 'A', l = 64, m = -14, n = 64; i < 157; i++ , m++) {
                        if (64 + m > 90) {
                            l += 1;
                            m = m - 26;
                        }
                        else if (64 + m == 90) {
                            l += 1;
                            m = 0;
                        }
                        if (i <= 39) {
                            k = String.fromCharCode(65 + m);
                        }
                        else {
                            k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                        }
                        if (i < this.growthRateLYSP.value && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                        }
                        else
                            if (i < this.growthRateLYSP.value && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                            }
                        if (i >= this.growthRateLYSP.value) {
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                            }
                            else
                                if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                    this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                                }
                                else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                    if (k == 'DR') {
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, '=round(Sum(DR' + (this.lastRowRoot1 + 1) + ':EC' + (this.lastRowRoot1 + 1) + ')+ ( Sum(DR' + (this.lastRowRoot1 + 1) + ':EC' + (this.lastRowRoot1 + 1) + ') * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                        m += 12;
                                    }
                                    else {
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, '=round(' + k + '' + (this.lastRowRoot1 + 1) + '+ (' + k + '' + (this.lastRowRoot1 + 1) + ' * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                    }
                                    j++;
                                }
                        }
                        else if (this.flexSheet.getCellData(3, i, this.treeItem) != 'Class' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Id' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Total') {
                            if (k == 'DR') {
                                m += 12;
                            }
                            this.flexSheet.setCellData(this.lastRowRoot1, i, 0);
                        }
                    }
                    this.nodes11[0].children.push({
                        id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    //  Applying Color for empty cell
                    this.flexSheet.applyCellsStyle({
                        backgroundColor: '#ffff99'
                    }, [new wjcGrid.CellRange(this.lastRowRoot1, 0, this.lastRowRoot1, this.growthRateLYSP.value - 1)]);
                    this.treeItem.treeModel.update();
                    this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
                }
                else if (this.nodeItem.data.id == 100) {
                    this.flexSheet.insertRows(this.lastRowRoot2, 1);
                    /*his.flexSheet.rows.insert(this.lastRowRoot2, new wjcGrid.Row());*/
                    this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    for (var i = 1, j = 2, k = 'A', l = 64, m = -14, n = 64; i < 157; i++ , m++) {
                        if (64 + m > 90) {
                            l += 1;
                            m = m - 26;
                        }
                        else if (64 + m == 90) {
                            l += 1;
                            m = 0;
                        }
                        if (i <= 39) {
                            k = String.fromCharCode(65 + m);
                        }
                        else {
                            k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                        }
                        if (i < this.growthRateLYSP.value && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                        }
                        else
                            if (i < this.growthRateLYSP.value && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                            }
                        if (i >= this.growthRateLYSP.value) {
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                            }
                            else
                                if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                                }
                                else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                    if (k == 'DR') {
                                        this.flexSheet.setCellData(this.lastRowRoot2, i, '=round(Sum(DR' + (this.lastRowRoot2 + 1) + ':EC' + (this.lastRowRoot2 + 1) + ')+ ( Sum(DR' + (this.lastRowRoot2 + 1) + ':EC' + (this.lastRowRoot2 + 1) + ') * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                        m += 12;
                                    }
                                    else {
                                        this.flexSheet.setCellData(this.lastRowRoot2, i, '=round(' + k + '' + (this.lastRowRoot2 + 1) + '+ (' + k + '' + (this.lastRowRoot2 + 1) + ' * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                    }
                                    j++;
                                }
                        }
                        else if (this.flexSheet.getCellData(3, i, this.treeItem) != 'Class' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Id' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Total') {
                            if (k == 'DR') {
                                m += 12;
                            }
                            this.flexSheet.setCellData(this.lastRowRoot2, i, 0);
                        }
                    }
                    this.nodes11[1].children.push({
                        id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    //  Applying Color for empty cell
                    this.flexSheet.applyCellsStyle({
                        backgroundColor: '#ffff99'
                    }, [new wjcGrid.CellRange(this.lastRowRoot2, 0, this.lastRowRoot2, this.growthRateLYSP.value - 1)]);
                    this.treeItem.treeModel.update();
                    this.lastRowRoot2 += 1;
                }
                this.edited = false;
                this.CategoryName = '';
                this._applyStyleForExpenceReportCss(this.formulaSheet1);
            }
            else if (this.EntryLastYear == 'ImmediatePrior') {
                if (this.nodeItem.data.id == 1) {
                    this.flexSheet.insertRows(this.lastRowRoot1, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot1, new wjcGrid.Row());
                    var minVal = 2, k = 'B', l = 65, m = 2
                    minVal = this.growthRate.value + 1;
                    if (minVal <= 26) {
                        m = minVal, l = 65;
                    } else {
                        m = (minVal % 26), l = 64 + Math.floor(minVal / 26);
                    }
                    if (Math.floor(minVal / 26) == 0) {
                        k = String.fromCharCode(64 + m);
                    }
                    else {
                        k = String.fromCharCode(l) + String.fromCharCode(64 + m);
                    }
                    for (var i = 1, j = 2; i < 157; i++) {
                        if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                        }
                        else if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                        }
                        else if (i >= minVal) {
                            if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot1, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot1 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot1 + 1) + ')');
                            }
                            else
                                if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                    this.flexSheet.setCellData(this.lastRowRoot1, i, this.flexSheet.getCellData((this.lastRowRoot1 - 1), i, this.treeItem));
                                }
                                else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                    if (k == 'EC') {
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, '=round(Sum(DR' + (this.lastRowRoot1 + 1) + ':EC' + (this.lastRowRoot1 + 1) + ')+ ( Sum(DR' + (this.lastRowRoot1 + 1) + ':EC' + (this.lastRowRoot1 + 1) + ') * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                        j++;
                                    }
                                    else {
                                        this.flexSheet.setCellData(this.lastRowRoot1, i, '=round(' + k + '' + (this.lastRowRoot1 + 1) + '+ (' + k + '' + (this.lastRowRoot1 + 1) + ' * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                        j++;
                                    }
                                    if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                        l += 1;
                                        m = m - 26;
                                    }
                                    else if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                        l += 1;
                                        m = 0;
                                    }
                                    if (k.length == 1 && k == 'Z') {
                                        m = 0;
                                    }
                                    if (k.length == 1 && k != 'Z') {
                                        k = String.fromCharCode(65 + m);
                                    }
                                    else {
                                        k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                    }
                                }
                            m++;
                        }
                        else if (this.flexSheet.getCellData(3, i, this.treeItem) != 'Class' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Id' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot1, i, 0);
                        }
                    }
                    this.nodes11[0].children.push({
                        id: this.nodes11[0].children.length == 0 ? 101 : this.nodes11[0].children[this.nodes11[0].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    //  Applying Color for empty cell
                    this.flexSheet.applyCellsStyle({
                        backgroundColor: '#ffff99'
                    }, [new wjcGrid.CellRange(this.lastRowRoot1, 0, this.lastRowRoot1, minVal - 1)]);
                    this.treeItem.treeModel.update();
                    this.lastRowRoot1 += 1; this.lastRowRoot2 += 1;
                }
                else if (this.nodeItem.data.id == 100) {
                    this.flexSheet.insertRows(this.lastRowRoot2, 1);
                    //this.flexSheet.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    this.formulaSheet1.rows.insert(this.lastRowRoot2, new wjcGrid.Row());
                    var minVal = 2, k = 'B', l = 65, m = 2
                    minVal = this.growthRate.value + 1;
                    if (minVal <= 26) {
                        m = minVal, l = 65;
                    } else {
                        m = (minVal % 26), l = 64 + Math.floor(minVal / 26);
                    }
                    if (Math.floor(minVal / 26) == 0) {
                        k = String.fromCharCode(64 + m);
                    }
                    else {
                        k = String.fromCharCode(l) + String.fromCharCode(64 + m);
                    }
                    for (var i = 1, j = 2; i < 157; i++) {
                        if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                        }
                        else
                            if (i < minVal && this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                            }
                            else if (i >= minVal) {
                                if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Total') {
                                    this.flexSheet.setCellData(this.lastRowRoot2, i, '=sum(' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 12) + (this.lastRowRoot2 + 1) + ':' + wjcGridSheet.FlexSheet.convertNumberToAlpha(i - 1) + (this.lastRowRoot2 + 1) + ')');
                                }
                                else
                                    if (this.flexSheet.getCellData(3, i, this.treeItem) == 'Class') {
                                        this.flexSheet.setCellData(this.lastRowRoot2, i, this.flexSheet.getCellData((this.lastRowRoot2 - 1), i, this.treeItem));
                                    }
                                    else if (!(this.flexSheet.getCellData(3, i, this.treeItem) == 'Id')) {
                                        if (k == 'EC') {
                                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=round(Sum(DR' + (this.lastRowRoot2 + 1) + ':EC' + (this.lastRowRoot2 + 1) + ')+ ( Sum(DR' + (this.lastRowRoot2 + 1) + ':EC' + (this.lastRowRoot2 + 1) + ') * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                            j++;
                                        }
                                        else {
                                            this.flexSheet.setCellData(this.lastRowRoot2, i, '=round(' + k + '' + (this.lastRowRoot2 + 1) + '+ (' + k + '' + (this.lastRowRoot2 + 1) + ' * ' + this.growthRateSheet.getCellData(j, 1, true) + ' / 100),0)');
                                            j++;
                                        }
                                        if (k.length == 2 && (k.charCodeAt(1) == 90 || m > 26)) {
                                            l += 1;
                                            m = m - 26;
                                        } else
                                            if (k.length == 2 && (k.charCodeAt(1) == 90 || m == 26)) {
                                                l += 1;
                                                m = 0;
                                            }
                                        if (k.length == 1 && k == 'Z') {
                                            m = 0;
                                        }
                                        if (k.length == 1 && k != 'Z') {
                                            k = String.fromCharCode(65 + m);
                                        }
                                        else {
                                            k = String.fromCharCode(l) + String.fromCharCode(65 + m);
                                        }
                                    }
                                m++;
                            }
                            else if (this.flexSheet.getCellData(3, i, this.treeItem) != 'Class' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Id' && this.flexSheet.getCellData(3, i, this.treeItem) != 'Total') {
                                this.flexSheet.setCellData(this.lastRowRoot2, i, 0);
                            }
                    }
                    this.nodes11[1].children.push({
                        id: this.nodes11[1].children.length == 0 ? 1001 : this.nodes11[1].children[this.nodes11[1].children.length - 1].id + 1,
                        name: newCat,
                        ischecked: false
                    });
                    //  Applying Color for empty cell cate2
                    this.flexSheet.applyCellsStyle({
                        backgroundColor: '#ffff99'
                    }, [new wjcGrid.CellRange(this.lastRowRoot2, 0, this.lastRowRoot2, minVal - 1)]);
                    this.treeItem.treeModel.update();
                    this.lastRowRoot2 += 1;

                }

                this.edited = false;
                this.CategoryName = '';
                this._applyStyleForExpenceReportCss(this.formulaSheet1);
            }
        }
        if (this.nodeItem.data.id == 1) {
            this.saveAsync(this.flexSheet, this.lastRowRoot1 - 1, newCat);
        }
        else if (this.nodeItem.data.id == 100) {
            this.saveAsync(this.flexSheet, this.lastRowRoot2 - 1, newCat);
        }
        this.nodesPerCat.forEach(obj => {
            obj.children.forEach(objInner => {
                objInner.ischecked = false;
            });
        });
        setTimeout(() => { this.resetColWidth(this.flexSheet); }, 200)
        this.treeItem.treeModel.expandAll();
        this.expandType = true;

        this.expandtreeDetails(this.treeItem.treeModel, this.nodeItem, $event);
        this.flexSheet.scrollPosition = new wjcCore.Point(scrollposition.x);
    }
    removeNode(tree: any, node: any) {
        let scrollposition = new wjcCore.Point(this.flexSheet.scrollPosition.x);
        if (node.parent.data.id == 1) {
            if (this.flexSheet.getCellData(node.index + 5, 14, tree)) {
                let IdIncremental = 15;
                for (var i = 14; i < 156; i += IdIncremental) {
                    if (i == 134) {
                        IdIncremental = 3;
                    }
                    this.deleteCat.push(parseInt(this.flexSheet.getCellData(node.index + 5, i, tree)));
                }
                this.budgetService.deleteExcelCat(this.deleteCat)
                    .subscribe(response => {
                        this.deleteCat = [];
                    });
            };
            this.flexSheet.deleteRows(node.index + 5, 1);
            this.formulaSheet1.rows.removeAt(node.index + 5);
            this.nodes11[0].children.splice(node.index, 1);
            tree.treeModel.update();
            this._applyStyleForExpenceReportCss(this.formulaSheet1);
            this.lastRowRoot1 += -1; this.lastRowRoot2 += -1;
        }
        else if (node.parent.data.id == 100) {
            if (this.flexSheet.getCellData(node.index + this.lastRowRoot1 + 1, 14, tree)) {
                let IdIncremental = 15;
                for (var i = 14; i < 156; i += IdIncremental) {
                    if (i == 134) {
                        IdIncremental = 3;
                    }
                    this.deleteCat.push(parseInt(this.flexSheet.getCellData(node.index + this.lastRowRoot1 + 1, i, tree)));
                }
                this.budgetService.deleteExcelCat(this.deleteCat)
                    .subscribe(response => {
                        this.deleteCat = [];
                    });
            };
            this.flexSheet.deleteRows(node.index + this.lastRowRoot1 + 1);
            // this.flexSheet.rows.removeAt(node.index + this.lastRowRoot1 + 1);
            this.formulaSheet1.rows.removeAt(node.index + this.lastRowRoot1 + 1);
            this.nodes11[1].children.splice(node.index, 1);
            tree.treeModel.update();
            //this._applyStyleForExpenceReport( this.flexSheet);
            this._applyStyleForExpenceReportCss(this.formulaSheet1);
            this.lastRowRoot2 += -1;
        }
        setTimeout(() => { this.resetColWidth(this.flexSheet); }, 200)
        this.flexSheet.scrollPosition = new wjcCore.Point(scrollposition.x);
    }

    constructor(private budgetService: BudgetService, private route: Router, private contextMenuService: ContextMenuService, private modalService: BsModalService) {

    }

    formulaSheetCss(formulaSheetCss: wjcGridSheet.FlexSheet) {
        this._applyStyleForExpenceReportCss(formulaSheetCss);

    }
    formulaSheetInit(formulaSheet: wjcGridSheet.FlexSheet) {
        let inputChild = 0;
        let inputArr;
        this.budgetService.getExcel()
            .subscribe(response => {
                let parseData = JSON.parse(response);
                inputArr = parseData;
                let self = this;

                if (formulaSheet) {
                    //var host = formulaSheet.hostElement;
                    formulaSheet.hostElement.addEventListener('mousedown', function (e) {
                        formulaSheet.applyCellsStyle({
                            backgroundColor: ''
                        });
                    });
                    formulaSheet.beginningEdit.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                        let data = formulaSheet.getCellData(args.range.row, args.range.col, true);
                        let class1 = this.lastRowRoot1 - this.nodes11[0].children.length - 1;
                        let class2 = this.lastRowRoot2 - this.nodes11[1].children.length - 1;
                        if (args.range.row == class1 || args.range.row == class2) {
                            args.cancel = true;
                        }
                        if (data.substr(0, 1) == "=" || data == "") {
                            args.cancel = true;
                        }
                        else if (args.range.row == 0 || args.range.row == 1 || args.range.row == 3) {
                            args.cancel = true;
                        }

                    });
                    formulaSheet.cellEditEnded.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                        let selection = args.range;
                        if (selection.isValid) {
                            //   self.currentCellData = formulaSheet.getCellData(selection.row, selection.col, true);
                            let value = formulaSheet.getCellData(args.range.row, args.range.col, true);
                            if (value.substr(0, 1) != "=") {
                                value = Math.round(formulaSheet.getCellData(args.range.row, args.range.col, false));
                                formulaSheet.setCellData(args.range.row, args.range.col, value);
                                formulaSheet.autoSizingColumn;
                                // formulaSheet.finishEditing(true);
                                //formulaSheet.columns.isFrozen(args.range.col);
                                //   formulaSheet.columns.frozen;
                                // sender.autoSizeColumn(args.range.col);
                            }
                            //setTimeout(() => { this.resetColWidth(this.flexSheet); }, 0);
                        }
                    });



                    formulaSheet.scrollPositionChanged.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                        let VerticalScrollPos = sender.scrollPosition.y;
                        let scrollposition = new wjcCore.Point(sender.scrollPosition.y);
                        this.formulaSheet1.scrollPosition = formulaSheet.scrollPosition.clone();
                        var elmnt = this.catDiv.nativeElement;
                        elmnt.scrollTop = -VerticalScrollPos;
                    });

                    formulaSheet.deferUpdate(() => {
                        self._generateExpenceReport(formulaSheet, inputChild, inputArr);
                    });
                }
            });
        //   this.resetColWidth(this.flexSheet);
    }

    // Set content for the use case template sheet.
    private _generateExpenceReport(flexSheet: wjcGridSheet.FlexSheet, data: Number, inputArr: Array<Array<Number>>) {
        flexSheet.setCellData(0, 1, 'Historical Data');
        flexSheet.setCellData(0, 76, 'Base Year');
        flexSheet.setCellData(0, 91, 'Forecast Budget');
        flexSheet.setCellData(1, 1, '2012');
        flexSheet.setCellData(1, 16, '2013');
        flexSheet.setCellData(1, 31, '2014');
        flexSheet.setCellData(1, 46, '2015');
        flexSheet.setCellData(1, 61, '2016');
        flexSheet.setCellData(1, 76, '2017');
        flexSheet.setCellData(1, 91, '2018');
        flexSheet.setCellData(1, 106, '2019');
        flexSheet.setCellData(1, 121, '2020');
        flexSheet.setCellData(1, 136, '2021');
        flexSheet.setCellData(1, 139, '2022');
        flexSheet.setCellData(1, 142, '2023');
        flexSheet.setCellData(1, 145, '2024');
        flexSheet.setCellData(1, 148, '2025');
        flexSheet.setCellData(1, 151, '2026');
        flexSheet.setCellData(1, 154, '2027');
        flexSheet.setCellData(3, 1, 'Jan');
        flexSheet.setCellData(3, 2, 'Feb');
        flexSheet.setCellData(3, 3, 'March');
        flexSheet.setCellData(3, 4, 'April');
        flexSheet.setCellData(3, 5, 'May');
        flexSheet.setCellData(3, 6, 'June');
        flexSheet.setCellData(3, 7, 'July');
        flexSheet.setCellData(3, 8, 'Aug');
        flexSheet.setCellData(3, 9, 'Sept');
        flexSheet.setCellData(3, 10, 'Oct');
        flexSheet.setCellData(3, 11, 'Nov');
        flexSheet.setCellData(3, 12, 'Dec');
        flexSheet.setCellData(3, 13, 'Total');
        flexSheet.setCellData(3, 14, 'Id');
        flexSheet.setCellData(3, 15, 'Class');
        flexSheet.setCellData(3, 16, 'Jan');
        flexSheet.setCellData(3, 17, 'Feb');
        flexSheet.setCellData(3, 18, 'March');
        flexSheet.setCellData(3, 19, 'April');
        flexSheet.setCellData(3, 20, 'May');
        flexSheet.setCellData(3, 21, 'June');
        flexSheet.setCellData(3, 22, 'July');
        flexSheet.setCellData(3, 23, 'Aug');
        flexSheet.setCellData(3, 24, 'Sept');
        flexSheet.setCellData(3, 25, 'Oct');
        flexSheet.setCellData(3, 26, 'Nov');
        flexSheet.setCellData(3, 27, 'Dec');
        flexSheet.setCellData(3, 28, 'Total');
        flexSheet.setCellData(3, 29, 'Id');
        flexSheet.setCellData(3, 30, 'Class');
        flexSheet.setCellData(3, 31, 'Jan');
        flexSheet.setCellData(3, 32, 'Feb');
        flexSheet.setCellData(3, 33, 'March');
        flexSheet.setCellData(3, 34, 'April');
        flexSheet.setCellData(3, 35, 'May');
        flexSheet.setCellData(3, 36, 'June');
        flexSheet.setCellData(3, 37, 'July');
        flexSheet.setCellData(3, 38, 'Aug');
        flexSheet.setCellData(3, 39, 'Sept');
        flexSheet.setCellData(3, 40, 'Oct');
        flexSheet.setCellData(3, 41, 'Nov');
        flexSheet.setCellData(3, 42, 'Dec');
        flexSheet.setCellData(3, 43, 'Total');
        flexSheet.setCellData(3, 44, 'Id');
        flexSheet.setCellData(3, 45, 'Class');
        flexSheet.setCellData(3, 46, 'Jan');
        flexSheet.setCellData(3, 47, 'Feb');
        flexSheet.setCellData(3, 48, 'March');
        flexSheet.setCellData(3, 49, 'April');
        flexSheet.setCellData(3, 50, 'May');
        flexSheet.setCellData(3, 51, 'June');
        flexSheet.setCellData(3, 52, 'July');
        flexSheet.setCellData(3, 53, 'Aug');
        flexSheet.setCellData(3, 54, 'Sept');
        flexSheet.setCellData(3, 55, 'Oct');
        flexSheet.setCellData(3, 56, 'Nov');
        flexSheet.setCellData(3, 57, 'Dec');
        flexSheet.setCellData(3, 58, 'Total');
        flexSheet.setCellData(3, 59, 'Id');
        flexSheet.setCellData(3, 60, 'Class');
        flexSheet.setCellData(3, 61, 'Jan');
        flexSheet.setCellData(3, 62, 'Feb');
        flexSheet.setCellData(3, 63, 'March');
        flexSheet.setCellData(3, 64, 'April');
        flexSheet.setCellData(3, 65, 'May');
        flexSheet.setCellData(3, 66, 'June');
        flexSheet.setCellData(3, 67, 'July');
        flexSheet.setCellData(3, 68, 'Aug');
        flexSheet.setCellData(3, 69, 'Sept');
        flexSheet.setCellData(3, 70, 'Oct');
        flexSheet.setCellData(3, 71, 'Nov');
        flexSheet.setCellData(3, 72, 'Dec');
        flexSheet.setCellData(3, 73, 'Total');
        flexSheet.setCellData(3, 74, 'Id');
        flexSheet.setCellData(3, 75, 'Class');
        flexSheet.setCellData(3, 76, 'Jan');
        flexSheet.setCellData(3, 77, 'Feb');
        flexSheet.setCellData(3, 78, 'March');
        flexSheet.setCellData(3, 79, 'April');
        flexSheet.setCellData(3, 80, 'May');
        flexSheet.setCellData(3, 81, 'June');
        flexSheet.setCellData(3, 82, 'July');
        flexSheet.setCellData(3, 83, 'Aug');
        flexSheet.setCellData(3, 84, 'Sept');
        flexSheet.setCellData(3, 85, 'Oct');
        flexSheet.setCellData(3, 86, 'Nov');
        flexSheet.setCellData(3, 87, 'Dec');
        flexSheet.setCellData(3, 88, 'Total');
        flexSheet.setCellData(3, 89, 'Id');
        flexSheet.setCellData(3, 90, 'Class');
        flexSheet.setCellData(3, 91, 'Jan');
        flexSheet.setCellData(3, 92, 'Feb');
        flexSheet.setCellData(3, 93, 'March');
        flexSheet.setCellData(3, 94, 'April');
        flexSheet.setCellData(3, 95, 'May');
        flexSheet.setCellData(3, 96, 'June');
        flexSheet.setCellData(3, 97, 'July');
        flexSheet.setCellData(3, 98, 'Aug');
        flexSheet.setCellData(3, 99, 'Sept');
        flexSheet.setCellData(3, 100, 'Oct');
        flexSheet.setCellData(3, 101, 'Nov');
        flexSheet.setCellData(3, 102, 'Dec');
        flexSheet.setCellData(3, 103, 'Total');
        flexSheet.setCellData(3, 104, 'Id');
        flexSheet.setCellData(3, 105, 'Class')
        flexSheet.setCellData(3, 106, 'Jan');
        flexSheet.setCellData(3, 107, 'Feb');
        flexSheet.setCellData(3, 108, 'March');
        flexSheet.setCellData(3, 109, 'April');
        flexSheet.setCellData(3, 110, 'May');
        flexSheet.setCellData(3, 111, 'June');
        flexSheet.setCellData(3, 112, 'July');
        flexSheet.setCellData(3, 113, 'Aug');
        flexSheet.setCellData(3, 114, 'Sept');
        flexSheet.setCellData(3, 115, 'Oct');
        flexSheet.setCellData(3, 116, 'Nov');
        flexSheet.setCellData(3, 117, 'Dec');
        flexSheet.setCellData(3, 118, 'Total');
        flexSheet.setCellData(3, 119, 'Id');
        flexSheet.setCellData(3, 120, 'Class')
        flexSheet.setCellData(3, 121, 'Jan');
        flexSheet.setCellData(3, 122, 'Feb');
        flexSheet.setCellData(3, 123, 'March');
        flexSheet.setCellData(3, 124, 'April');
        flexSheet.setCellData(3, 125, 'May');
        flexSheet.setCellData(3, 126, 'June');
        flexSheet.setCellData(3, 127, 'July');
        flexSheet.setCellData(3, 128, 'Aug');
        flexSheet.setCellData(3, 129, 'Sept');
        flexSheet.setCellData(3, 130, 'Oct');
        flexSheet.setCellData(3, 131, 'Nov');
        flexSheet.setCellData(3, 132, 'Dec');
        flexSheet.setCellData(3, 133, 'Total');
        flexSheet.setCellData(3, 134, 'Id');
        flexSheet.setCellData(3, 135, 'Class');
        flexSheet.setCellData(3, 136, '');
        flexSheet.setCellData(3, 137, 'Id');
        flexSheet.setCellData(3, 138, 'Class');
        flexSheet.setCellData(3, 139, '');
        flexSheet.setCellData(3, 140, 'Id');
        flexSheet.setCellData(3, 141, 'Class');
        flexSheet.setCellData(3, 142, '');
        flexSheet.setCellData(3, 143, 'Id');
        flexSheet.setCellData(3, 144, 'Class');
        flexSheet.setCellData(3, 145, '');
        flexSheet.setCellData(3, 146, 'Id');
        flexSheet.setCellData(3, 147, 'Class');
        flexSheet.setCellData(3, 148, '');
        flexSheet.setCellData(3, 149, 'Id');
        flexSheet.setCellData(3, 150, 'Class');
        flexSheet.setCellData(3, 151, '');
        flexSheet.setCellData(3, 152, 'Id');
        flexSheet.setCellData(3, 153, 'Class');
        flexSheet.setCellData(3, 154, '');
        flexSheet.setCellData(3, 155, 'Id');
        flexSheet.setCellData(3, 156, 'Class');

        this._setExpenseData(flexSheet, data, inputArr);

        this._applyStyleForExpenceReport(flexSheet);
        this._applyStyleForExpenceReportCss(this.formulaSheet1);
    }

    // set expense detail data for the use case template sheet.
    private _setExpenseData(flexSheet: wjcGridSheet.FlexSheet, data: Number, inputArr: Array<Array<Number>>) {
        let
            colIndex,
            rowIndex,
            value,
            currValue, maxRow,
            jan = 1, feb = 2, mar = 3, apl = 4, may = 5, jun = 6, jly = 7, aug = 8, sep = 9, oct = 10, nov = 11, dec = 12, total = 13, id = 14, Class = 15;
        maxRow = inputArr[0];

        for (rowIndex = 4; rowIndex < maxRow.yearViewModel.length + 4; rowIndex++) {
            currValue = 0;
            let arrayData;
            for (var i = 0; i < inputArr.length; i++) {
                if (i <= 8) {
                    value = inputArr[i];
                    arrayData = value.yearViewModel;
                    flexSheet.setCellData(rowIndex, id + currValue, value.yearViewModel[rowIndex - 4]['id'], true);
                    flexSheet.setCellData(rowIndex, jan + currValue, value.yearViewModel[rowIndex - 4]['jan'], true);
                    flexSheet.setCellData(rowIndex, feb + currValue, value.yearViewModel[rowIndex - 4]['feb'], true);
                    flexSheet.setCellData(rowIndex, mar + currValue, value.yearViewModel[rowIndex - 4]['march'], true);
                    flexSheet.setCellData(rowIndex, apl + currValue, value.yearViewModel[rowIndex - 4]['april'], true);
                    flexSheet.setCellData(rowIndex, may + currValue, value.yearViewModel[rowIndex - 4]['may'], true);
                    flexSheet.setCellData(rowIndex, jun + currValue, value.yearViewModel[rowIndex - 4]['june'], true);
                    flexSheet.setCellData(rowIndex, jly + currValue, value.yearViewModel[rowIndex - 4]['july'], true);
                    flexSheet.setCellData(rowIndex, aug + currValue, value.yearViewModel[rowIndex - 4]['aug'], true);
                    flexSheet.setCellData(rowIndex, sep + currValue, value.yearViewModel[rowIndex - 4]['sept'], true);
                    flexSheet.setCellData(rowIndex, oct + currValue, value.yearViewModel[rowIndex - 4]['oct'], true);
                    flexSheet.setCellData(rowIndex, nov + currValue, value.yearViewModel[rowIndex - 4]['nov'], true);
                    flexSheet.setCellData(rowIndex, dec + currValue, value.yearViewModel[rowIndex - 4]['dec'], true);
                    flexSheet.setCellData(rowIndex, Class + currValue, value.yearViewModel[rowIndex - 4]['class'], true);
                    //getvalue in var
                    flexSheet.setCellData(rowIndex, total + currValue, value.yearViewModel[rowIndex - 4]['total'], true);
                    currValue += 15;
                }
                else {
                    value = inputArr[i];
                    flexSheet.setCellData(rowIndex, feb + currValue, value.yearViewModel[rowIndex - 4]['id'], true);
                    flexSheet.setCellData(rowIndex, jan + currValue, value.yearViewModel[rowIndex - 4]['jan'], true);
                    flexSheet.setCellData(rowIndex, mar + currValue, value.yearViewModel[rowIndex - 4]['class'], true);
                    currValue += 3;
                }

            }

        }
    }

    // Apply styles for the use case template sheet.
    private _applyStyleForExpenceReport(flexSheet: wjcGridSheet.FlexSheet) {
        //flexSheet.columns.defaultSize = 50;
        flexSheet.frozenRows = 4;
        flexSheet.applyCellsStyle({
            fontSize: '16px',
            fontWeight: 'bold',
            color: '#696964',
        }, [new wjcGrid.CellRange(1, 1, 1, 156)]);
        flexSheet.mergeRange(new wjcGrid.CellRange(0, 1, 0, 75));
        flexSheet.applyCellsStyle({
            backgroundColor: '#caffca',
        }, [new wjcGrid.CellRange(0, 1, 0, 75)]);
        flexSheet.mergeRange(new wjcGrid.CellRange(0, 76, 0, 90));
        flexSheet.applyCellsStyle({
            backgroundColor: '#f3c975',
        }, [new wjcGrid.CellRange(0, 76, 0, 90)]);
        flexSheet.mergeRange(new wjcGrid.CellRange(0, 91, 0, 156));
        flexSheet.applyCellsStyle({
            backgroundColor: '#ffb3b3',
        }, [new wjcGrid.CellRange(0, 91, 0, 156)]);
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 1, 1, 15));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 16, 1, 30));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 31, 1, 45));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 46, 1, 60));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 61, 1, 75));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 76, 1, 90));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 91, 1, 105));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 106, 1, 120));
        flexSheet.mergeRange(new wjcGrid.CellRange(1, 121, 1, 135));
        flexSheet.applyCellsStyle({
            backgroundColor: '#ffe6cd',
        }, [new wjcGrid.CellRange(1, 1, 1, 15), new wjcGrid.CellRange(1, 31, 1, 45), new wjcGrid.CellRange(1, 61, 1, 75), new wjcGrid.CellRange(1, 91, 1, 105), new wjcGrid.CellRange(1, 121, 1, 135), new wjcGrid.CellRange(1, 139, 1, 141), new wjcGrid.CellRange(1, 145, 1, 147), new wjcGrid.CellRange(1, 151, 1, 153)]);
        flexSheet.applyCellsStyle({
            backgroundColor: '#ffffca',
        }, [new wjcGrid.CellRange(1, 16, 1, 30), new wjcGrid.CellRange(1, 46, 1, 60), new wjcGrid.CellRange(1, 76, 1, 90), new wjcGrid.CellRange(1, 106, 1, 120), new wjcGrid.CellRange(1, 136, 1, 138), new wjcGrid.CellRange(1, 142, 1, 144), new wjcGrid.CellRange(1, 148, 1, 150), new wjcGrid.CellRange(1, 154, 1, 156)]);
        flexSheet.applyCellsStyle({
            fontWeight: 'bold',
            backgroundColor: '#dce3ee',
        }, [new wjcGrid.CellRange(3, 0, 3, 156)]);
        flexSheet.applyCellsStyle({
            textAlign: 'center'
        }, [new wjcGrid.CellRange(3, 1, 3, 156)]);



        let class1 = this.lastRowRoot1 - this.nodes11[0].children.length - 1;
        let class2 = this.lastRowRoot2 - this.nodes11[1].children.length - 1;

        for (var i = 4; i < flexSheet.rows.length; i++) {
            for (var j = 0; j < flexSheet.columns.length; j++) {
                let value = flexSheet.getCellData(i, j, true);
                if (!((i == class1) || (i == class2))) {
                    if (value.substr(0, 1) != "=" && value !== "") {
                        flexSheet.applyCellsStyle({
                            backgroundColor: '#ffff99'
                        }, [new wjcGrid.CellRange(i, j)]);
                    }
                }
            }
        }

        flexSheet.rowHeaders.columns.removeAt(0);
        flexSheet.columnHeaders.rows.removeAt(0);
        this.resetColWidth(flexSheet);
        flexSheet.rows[2].height = 0;

    }

    // Styles for Main Grid
    private _applyStyleForExpenceReportCss(flexSheet: wjcGridSheet.FlexSheet) {
        flexSheet.columns[0].width = 109;
        flexSheet.rows[2].height = 0;
        let minVal = 4; let maxVal = 4; let alternet = 0;
        for (var i = 0; i < 500; i++) {
            if (alternet == 0) {
                flexSheet.applyCellsStyle({
                    backgroundColor: '#e8e8fa'
                }, [new wjcGrid.CellRange(maxVal, 0, maxVal, 0)]);

                alternet = 1;
            }
            else {
                flexSheet.applyCellsStyle({
                    backgroundColor: ''
                }, [new wjcGrid.CellRange(minVal, 0, minVal, 0)]);
                alternet = 0;
            }
            minVal += 1; maxVal += 1;
        }

        flexSheet.rowHeaders.columns.removeAt(0);
        flexSheet.columnHeaders.rows.removeAt(0);

    }



    saveStatic(formulaSheet: wjcGridSheet.FlexSheet) {
        let flexSheet = formulaSheet;
        let rowIndex,
            colIndex,
            value,
            currValue = 0, currNode = 0, parNode = 0,
            jan = 1, feb = 2, mar = 3, apl = 4, may = 5, jun = 6, jly = 7, aug = 8, sep = 9, oct = 10, nov = 11, dec = 12, total = 13, Class = 15, id = 14, yearVal = 2012;
        for (colIndex = 0; colIndex < 16; colIndex++) {
            this.yearData = new Array<YearInfo>();
            if (colIndex <= 8) {
                for (rowIndex = 4; rowIndex < formulaSheet.rows.length; rowIndex++) {
                    if (formulaSheet.getCellValue(1, jan + currValue)) {
                        yearVal = formulaSheet.getCellValue(1, jan + currValue);
                    }
                    this.year.id = formulaSheet.getCellValue(rowIndex, id + currValue);
                    this.year.jan = formulaSheet.getCellData(rowIndex, jan + currValue, false);
                    this.year.feb = formulaSheet.getCellData(rowIndex, feb + currValue, false);
                    this.year.march = formulaSheet.getCellData(rowIndex, mar + currValue, false);
                    this.year.april = formulaSheet.getCellData(rowIndex, apl + currValue, false);
                    this.year.may = formulaSheet.getCellData(rowIndex, may + currValue, false);
                    this.year.june = formulaSheet.getCellData(rowIndex, jun + currValue, false);
                    this.year.july = formulaSheet.getCellData(rowIndex, jly + currValue, false);
                    this.year.aug = formulaSheet.getCellData(rowIndex, aug + currValue, false);
                    this.year.sept = formulaSheet.getCellData(rowIndex, sep + currValue, false);
                    this.year.oct = formulaSheet.getCellData(rowIndex, oct + currValue, false);
                    this.year.nov = formulaSheet.getCellData(rowIndex, nov + currValue, false);
                    this.year.dec = formulaSheet.getCellData(rowIndex, dec + currValue, false);
                    this.year.Class = formulaSheet.getCellValue(rowIndex, Class + currValue, false);
                    this.year.total = formulaSheet.getCellData(rowIndex, total + currValue, false);
                    this.year.year = yearVal;
                    if (rowIndex == 4) {
                        parNode = 0;
                        this.year.category = this.nodes11[parNode].name;
                        currNode = -1;
                    }
                    else if ((this.nodes11[parNode].children.length == currNode)) {
                        parNode += 1;
                        this.year.category = this.nodes11[parNode].name;
                        currNode = -1;

                    }
                    else {
                        this.year.category = this.nodes11[parNode].children[currNode].name;
                    }

                    this.yearData.push(this.year);
                    this.year = new YearInfo();
                    currNode += 1;
                }
                currValue += 15;
            } else {
                for (rowIndex = 4; rowIndex < formulaSheet.rows.length; rowIndex++) {
                    if (formulaSheet.getCellValue(1, jan + currValue)) {
                        yearVal = formulaSheet.getCellValue(1, jan + currValue);
                    }
                    this.year.id = formulaSheet.getCellValue(rowIndex, feb + currValue);
                    this.year.jan = formulaSheet.getCellData(rowIndex, jan + currValue, false);
                    this.year.feb = "";
                    this.year.march = "";
                    this.year.april = "";
                    this.year.may = "";
                    this.year.june = "";
                    this.year.july = "";
                    this.year.aug = "";
                    this.year.sept = "";
                    this.year.oct = "";
                    this.year.nov = "";
                    this.year.dec = "";
                    this.year.Class = formulaSheet.getCellValue(rowIndex, mar + currValue, false);
                    this.year.total = "";
                    this.year.year = yearVal;
                    if (rowIndex == 4) {
                        parNode = 0;
                        this.year.category = this.nodes11[parNode].name;
                        currNode = -1;
                    }
                    else if ((this.nodes11[parNode].children.length == currNode)) {
                        parNode += 1;
                        this.year.category = this.nodes11[parNode].name;
                        currNode = -1;

                    }
                    else {
                        this.year.category = this.nodes11[parNode].children[currNode].name;
                    }

                    this.yearData.push(this.year);
                    this.year = new YearInfo();
                    currNode += 1;
                }
                currValue += 3;
            }

            this.yearDataArr.push(this.yearData);
        }

        if (this.yearDataArr) {
            this.budgetService.saveExcel2(this.yearDataArr)
                .subscribe(response => {
                    let inputChild = JSON.parse(response)
                    this.budgetService._staticData = inputChild;
                    setTimeout(() => { location.reload(); }, 250);
                });
        }
    }
    openPopup(tree: any, node: any): void {
        this.nodeItem = node;
        this.treeItem = tree;
        this.edited = true;
    }
    closePopup(): void {

        this.edited = false;
    }

    ShowGrowthRateCategory(): void {
        this.isGrowthRate = true;
        this.isPercent = false;
    }
    ShowPercentCategory(): void {
        this.isGrowthRate = false;
        this.isPercent = true;
    }

    HidePopupCategory(): void {
        this.isGrowthRate = false;
        this.isPercent = false;
    }




    growthRateSheetInit(formulaSheet: wjcGridSheet.FlexSheet, confirmtemp: TemplateRef<any>) {
        let inputChild = 0;
        let inputArr;
        this._applyStyleForGrowthRateCss(formulaSheet);
        let self = this;
        let value;
        if (formulaSheet) {
            []
            formulaSheet.cellEditEnded.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                if (args.range.row == 2) {
                    this.args = args;
                    this.ispercentconfirm = false;
                    this.isGrowthconfirm = true;
                    let value = formulaSheet.getCellData(args.range.row, args.range.col, true);
                    if (value == "") {
                        formulaSheet.setCellData(args.range.row, args.range.col, 0);
                    }
                    this.modalRef = this.modalService.show(confirmtemp);
                }

            });
            formulaSheet.selectionChanged.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                let selection = args.range;
                if (selection.isValid) {
                    //self.currentCellData = formulaSheet.getCellData(selection.row, selection.col, true);
                    value = formulaSheet.getCellData(selection.row, selection.col, true);
                }
            });

            formulaSheet.beginningEdit.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                if (args.range.col == 0 || args.range.row == 0) {
                    args.cancel = true;
                }

            });

            formulaSheet.deferUpdate(() => {
                self._generateGrowthRateReport(formulaSheet);
            });
        }
        //   });
    }

    // Set content for the use case template sheet.
    private _generateGrowthRateReport(flexSheet: wjcGridSheet.FlexSheet) {
        if (this.EntryLastYear == 'LastYearSamePeriod') {
            let year = parseInt(this.growthRateLYSP.name);
            let months = 'Jan';
            let rowCount = parseInt(this.growthRateLYSP.name) >= 2020 ? 9 - (parseInt(this.growthRateLYSP.name) - 2020) : 105 - ((parseInt(this.growthRateLYSP.name) - 2012) * 12)
            flexSheet.setCellData(0, 0, 'Period');
            flexSheet.setCellData(0, 1, 'Growth %');
            for (var i = 2; i < rowCount; i++) {
                while (this.growthRateSheet.rows.length != rowCount) {
                    if (this.growthRateSheet.rows.length > rowCount) {
                        this.growthRateSheet.deleteRows(this.growthRateSheet.rows.length - 1);
                    }
                    else if (this.growthRateSheet.rows.length < rowCount) {
                        this.growthRateSheet.insertRows(this.growthRateSheet.rows.length, 1);
                    }
                }
                this.growthRateSheet.scrollPosition = new wjcCore.Point(0);
                flexSheet.setCellData(i, 1, 1);
                if (i <= (rowCount - 8)) {
                    year = months == 'Jan' ? year + 1 : year;
                    flexSheet.setCellData(i, 0, months + "  " + year);
                    months = months == 'Jan' ? 'Feb' : months == 'Feb' ? 'March' : months == 'March' ? 'April' : months == 'April' ? 'May' : months == 'May' ? 'June' : months == 'June' ? 'July' : months == 'July' ? 'Aug' : months == 'Aug' ? 'Sept' : months == 'Sept' ? 'Oct' : months == 'Oct' ? 'Nov' : months == 'Nov' ? 'Dec' : 'Jan';
                }
                else {
                    year += 1;
                    flexSheet.setCellData(i, 0, year);
                }

            }
        } else if (this.EntryLastYear == 'ImmediatePrior') {

            let year = parseInt(this.growthRate.name.split(' ')[1]);
            let maxRow = 117 - this.growthRate.value - ((2012 - year) * 3);
            let yearRow = 109 - this.growthRate.value - ((2012 - year) * 3);
            let months = this.growthRate.name.split(' ')[0];
            flexSheet.setCellData(0, 0, 'Period');
            flexSheet.setCellData(0, 1, 'Growth %');
            for (var i = 2; i < maxRow; i++) {
                while (this.growthRateSheet.rows.length != maxRow) {
                    if (this.growthRateSheet.rows.length > maxRow) {
                        this.growthRateSheet.rows.removeAt(this.growthRateSheet.rows.length - 1);
                    }
                    else if (this.growthRateSheet.rows.length < maxRow) {
                        this.growthRateSheet.rows.insert(this.growthRateSheet.rows.length, new wjcGrid.Row());
                    }
                }
                flexSheet.setCellData(i, 1, 1);
                if (i <= yearRow) {
                    months = months == 'Jan' ? 'Feb' : months == 'Feb' ? 'March' : months == 'March' ? 'April' : months == 'April' ? 'May' : months == 'May' ? 'June' : months == 'June' ? 'July' : months == 'July' ? 'Aug' : months == 'Aug' ? 'Sept' : months == 'Sept' ? 'Oct' : months == 'Oct' ? 'Nov' : months == 'Nov' ? 'Dec' : 'Jan';
                    year = months == 'Jan' ? year + 1 : year;
                    flexSheet.setCellData(i, 0, months + "  " + year);
                }
                else {
                    year += 1;
                    flexSheet.setCellData(i, 0, year);
                }

            }
        }

    }
    changePriorPeriod() {
        this._generateGrowthRateReport(this.growthRateSheet);
    };
    private _applyStyleForGrowthRateCss(flexSheet: wjcGridSheet.FlexSheet) {
        flexSheet.rowHeaders.columns.removeAt(0);
        flexSheet.columnHeaders.rows.removeAt(0);
        flexSheet.rows[1].height = 0;
        flexSheet.isTabHolderVisible = false;
        flexSheet.frozenRows = 1;

    }
    private growthRateSheetDblclicked(formulaSheet, args) {
        let select = args.range;
        let gridrow = select.row;
        let gridcol = select.col;
        let currentcelldata = formulaSheet.getCellData(select.row, select.col, true)
        for (var i = 1; i < formulaSheet.rows.length - 2; i++) {
            formulaSheet.setCellData(gridrow + i, gridcol, currentcelldata);
        }
    }
    // for popup
    modalRef: BsModalRef;
    modalRef2: BsModalRef;
    confirm(): void {
        this.modalRef.hide();
        let growthargs = this.args;
        let percentArgs = this.percentArgs;
        if (this.isGrowthconfirm) { this.growthRateSheetDblclicked(this.growthRateSheet, growthargs); }
        if (this.ispercentconfirm) { this.growthRateSheetDblclicked(this.percentRateSheet, percentArgs); }
    }

    decline(): void {
        this.modalRef.hide();
    }
    message(): void {
        this.modalRef.hide();
    }
    //for percentSheet
    private _applyStyleForPercentRateCss(flexSheet: wjcGridSheet.FlexSheet) {
        flexSheet.rowHeaders.columns.removeAt(0);
        flexSheet.columnHeaders.rows.removeAt(0);
        flexSheet.rows[1].height = 0;
        flexSheet.isTabHolderVisible = false;
        flexSheet.frozenRows = 1;
    }
    private _generatePercentRateReport(flexSheet: wjcGridSheet.FlexSheet) {
        let year = 2017;
        let months = 'Jan';
        flexSheet.setCellData(0, 0, 'Period');
        flexSheet.setCellData(0, 1, '%');
        for (var i = 2; i < 45; i++) {
            if (this.percentRateSheet.rows.length > 45) {
                this.percentRateSheet.rows.removeAt(this.percentRateSheet.rows.length - 1);
            }
            else if (this.percentRateSheet.rows.length < 45) {
                this.percentRateSheet.rows.insert(this.percentRateSheet.rows.length, new wjcGrid.Row());
            }
            flexSheet.setCellData(i, 1, 1);
            if (i <= 37) {
                year = months == 'Jan' ? year + 1 : year;
                flexSheet.setCellData(i, 0, months + "  " + year);
                months = months == 'Jan' ? 'Feb' : months == 'Feb' ? 'March' : months == 'March' ? 'April' : months == 'April' ? 'May' : months == 'May' ? 'June' : months == 'June' ? 'July' : months == 'July' ? 'Aug' : months == 'Aug' ? 'Sept' : months == 'Sept' ? 'Oct' : months == 'Oct' ? 'Nov' : months == 'Nov' ? 'Dec' : 'Jan';
            }
            else {
                year += 1;
                flexSheet.setCellData(i, 0, year);
            }

        }
    }
    percentRateSheetInit(formulaSheet: wjcGridSheet.FlexSheet, confirmtemp: TemplateRef<any>) {
        let inputChild = 0;
        let inputArr;
        this._applyStyleForPercentRateCss(formulaSheet);
        let self = this;
        if (formulaSheet) {
            []
            formulaSheet.cellEditEnded.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                if (args.range.row == 2) {
                    this.percentArgs = args;
                    this.ispercentconfirm = true;
                    this.isGrowthconfirm = false;
                    let value = formulaSheet.getCellData(args.range.row, args.range.col, true);
                    if (value == "") {
                        formulaSheet.setCellData(args.range.row, args.range.col, 0);
                    }
                    this.modalRef = this.modalService.show(confirmtemp);
                }

            });
            formulaSheet.selectionChanged.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                let selection = args.range;
                if (selection.isValid) {
                    self.currentCellData = formulaSheet.getCellData(selection.row, selection.col, true);
                }
            });
            formulaSheet.beginningEdit.addHandler((sender: any, args: wjcGrid.CellRangeEventArgs) => {
                if (args.range.col == 0 || args.range.row == 0) {
                    args.cancel = true;
                }
            });
            formulaSheet.deferUpdate(() => {
                self._generatePercentRateReport(formulaSheet);
            });
        }

    }

    private ISvalidCat(newCat) {
        this.validCat = true;
        if (this.nodeItem.data.id == 1) {
            if (newCat.toUpperCase() == this.nodeItem.data.name.replace(' ', '').toUpperCase()) {
                this.modalRef = this.modalService.show(this.validatemessage);
                this.validCat = false;
            }
            else {
                for (var i = 0; i < this.nodeItem.data.children.length; i++) {
                    if (newCat.toUpperCase() == this.nodeItem.data.children[i].name.toUpperCase()) {
                        this.modalRef = this.modalService.show(this.validatemessage);
                        //this.validCat = false;
                        this.validCat = false;
                    }
                }
            }
        }

        if (this.nodeItem.data.id == 100) {
            this.validCat = true;
            if (newCat.toUpperCase() == this.nodeItem.data.name.replace(' ', '').toUpperCase()) {
                this.modalRef = this.modalService.show(this.validatemessage);
                this.validCat = false;
            }
            else {
                for (var i = 0; i < this.nodeItem.data.children.length; i++) {
                    if (newCat.toUpperCase() == this.nodeItem.data.children[i].name.toUpperCase()) {
                        this.modalRef = this.modalService.show(this.validatemessage);
                        this.validCat = false;
                    }
                }
            }

        }
    }

    private resetColWidth(flexSheet: any) {
        flexSheet.autoSizeColumns();
        flexSheet.columns[14].width = 0;
        flexSheet.columns[15].width = 0;
        flexSheet.columns[29].width = 0;
        flexSheet.columns[30].width = 0;
        flexSheet.columns[44].width = 0;
        flexSheet.columns[45].width = 0;
        flexSheet.columns[59].width = 0;
        flexSheet.columns[60].width = 0;
        flexSheet.columns[74].width = 0;
        flexSheet.columns[75].width = 0;
        flexSheet.columns[89].width = 0;
        flexSheet.columns[90].width = 0;
        flexSheet.columns[104].width = 0;
        flexSheet.columns[105].width = 0;
        flexSheet.columns[119].width = 0;
        flexSheet.columns[120].width = 0;
        flexSheet.columns[134].width = 0;
        flexSheet.columns[135].width = 0;
        flexSheet.columns[137].width = 0;
        flexSheet.columns[138].width = 0;
        flexSheet.columns[140].width = 0;
        flexSheet.columns[141].width = 0;
        flexSheet.columns[143].width = 0;
        flexSheet.columns[144].width = 0;
        flexSheet.columns[146].width = 0;
        flexSheet.columns[147].width = 0;
        flexSheet.columns[149].width = 0;
        flexSheet.columns[150].width = 0;
        flexSheet.columns[152].width = 0;
        flexSheet.columns[153].width = 0;
        flexSheet.columns[155].width = 0;
        flexSheet.columns[156].width = 0;
        flexSheet.columns[0].width = 0;
    }

    private expandtreeDetails(tree, node, $event) {
        if (this.expandType) {
            for (var i = 5; i < this.lastRowRoot1; i++) {
                this.formulaSheet1.rows[i].height = 28;
                this.flexSheet.rows[i].height = 28;
            }
            for (var i = this.lastRowRoot1 + 1; i < this.lastRowRoot2; i++) {
                this.formulaSheet1.rows[i].height = 28;
                this.flexSheet.rows[i].height = 28;
            }
            this.expandType = false;
        }
        else {
            if (node.data.id == 1) {
                for (var i = 5; i < this.lastRowRoot1; i++) {
                    if (tree.isExpanded(node)) {
                        this.formulaSheet1.rows[i].height = 0;
                        this.flexSheet.rows[i].height = 0;
                    } else {
                        this.formulaSheet1.rows[i].height = 28;
                        this.flexSheet.rows[i].height = 28;
                    }

                }
            } else if (node.data.id == 100) {
                for (var i = this.lastRowRoot1 + 1; i < this.lastRowRoot2; i++) {
                    if (tree.isExpanded(node)) {
                        this.formulaSheet1.rows[i].height = 0;
                        this.flexSheet.rows[i].height = 0;

                    } else {
                        this.formulaSheet1.rows[i].height = 28;
                        this.flexSheet.rows[i].height = 28;

                    }

                }
            }
            TREE_ACTIONS.TOGGLE_EXPANDED(tree, node, $event);
        }

    }
    saveAsync(formulaSheet: wjcGridSheet.FlexSheet, AddedRow: number, CatName: string) {
        let flexSheet = formulaSheet;
        let rowIndex,
            colIndex,
            value,
            currValue = 0, currNode = 0, parNode = 0,
            jan = 1, feb = 2, mar = 3, apl = 4, may = 5, jun = 6, jly = 7, aug = 8, sep = 9, oct = 10, nov = 11, dec = 12, total = 13, Class = 15, id = 14, yearVal = 2012;
        for (colIndex = 0; colIndex < 16; colIndex++) {
            this.yearData = new Array<YearInfo>();
            if (colIndex <= 8) {
                rowIndex = AddedRow;
                if (formulaSheet.getCellValue(1, jan + currValue)) {
                    yearVal = formulaSheet.getCellValue(1, jan + currValue);
                }
                this.year.id = formulaSheet.getCellValue(rowIndex, id + currValue);
                this.year.jan = formulaSheet.getCellData(rowIndex, jan + currValue, false);
                this.year.feb = formulaSheet.getCellData(rowIndex, feb + currValue, false);
                this.year.march = formulaSheet.getCellData(rowIndex, mar + currValue, false);
                this.year.april = formulaSheet.getCellData(rowIndex, apl + currValue, false);
                this.year.may = formulaSheet.getCellData(rowIndex, may + currValue, false);
                this.year.june = formulaSheet.getCellData(rowIndex, jun + currValue, false);
                this.year.july = formulaSheet.getCellData(rowIndex, jly + currValue, false);
                this.year.aug = formulaSheet.getCellData(rowIndex, aug + currValue, false);
                this.year.sept = formulaSheet.getCellData(rowIndex, sep + currValue, false);
                this.year.oct = formulaSheet.getCellData(rowIndex, oct + currValue, false);
                this.year.nov = formulaSheet.getCellData(rowIndex, nov + currValue, false);
                this.year.dec = formulaSheet.getCellData(rowIndex, dec + currValue, false);
                this.year.Class = formulaSheet.getCellValue(rowIndex, Class + currValue, false);
                this.year.total = formulaSheet.getCellData(rowIndex, total + currValue, false);
                this.year.year = yearVal;
                this.year.category = CatName;

                this.yearData.push(this.year);
                this.year = new YearInfo();
                currNode += 1;
                currValue += 15;
            } else {
                rowIndex = AddedRow;
                if (formulaSheet.getCellValue(1, jan + currValue)) {
                    yearVal = formulaSheet.getCellValue(1, jan + currValue);
                }
                this.year.id = formulaSheet.getCellValue(rowIndex, feb + currValue);
                this.year.jan = formulaSheet.getCellData(rowIndex, jan + currValue, false);
                this.year.feb = "";
                this.year.march = "";
                this.year.april = "";
                this.year.may = "";
                this.year.june = "";
                this.year.july = "";
                this.year.aug = "";
                this.year.sept = "";
                this.year.oct = "";
                this.year.nov = "";
                this.year.dec = "";
                this.year.total = "";
                this.year.Class = formulaSheet.getCellValue(rowIndex, mar + currValue, false);
                this.year.year = yearVal;
                this.year.category = CatName;

                this.yearData.push(this.year);
                this.year = new YearInfo();
                currNode += 1;
                currValue += 3;
            }

            this.yearDataArr.push(this.yearData);
        }

        if (this.yearDataArr) {
            this.budgetService.saveExcel2(this.yearDataArr)
                .subscribe(response => {
                    let inputChild = JSON.parse(response)
                    this.yearDataArr = new Array<Array<YearInfo>>();
                    this._setAsyncSavedData(formulaSheet, AddedRow, inputChild);
                });
        }
    }

    private _setAsyncSavedData(flexSheet: wjcGridSheet.FlexSheet, rowAddedNo: number, inputArr: Array<Array<Number>>) {
        let
            colIndex,
            rowIndex,
            value,
            currValue, maxRow,
            feb = 2, mar = 3, Class = 15, id = 14;
        maxRow = inputArr[0];

        rowIndex = rowAddedNo;
        currValue = 0;

        for (var i = 0; i < inputArr.length; i++) {
            if (i <= 8) {
                value = inputArr[i];
                flexSheet.setCellData(rowIndex, id + currValue, value[0]['id'], true);
                flexSheet.setCellData(rowIndex, Class + currValue, value[0]['class'], true);
                currValue += 15;
            }
            else {
                value = inputArr[i];
                flexSheet.setCellData(rowIndex, feb + currValue, value[0]['id'], true);
                flexSheet.setCellData(rowIndex, mar + currValue, value[0]['class'], true);
                currValue += 3;
            }

        }

    }
}


