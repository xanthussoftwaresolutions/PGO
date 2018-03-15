import { Injectable } from '@angular/core';
import { Http, Response, Headers, RequestOptions } from '@angular/http';
import { Year } from '../_models/index';
import { Observable } from 'rxjs/Observable';
import "rxjs/Rx";
import * as wjcGridSheet from 'wijmo/wijmo.grid.sheet';

@Injectable()
export class BudgetService {
    
    public _saveExcelUrl: string = '/Budget/SaveExcelUrl/';
    public _saveExcel2Url: string = '/Budget/SaveExcel2Url/';
    public _getExcelUrl: string = '/Budget/GetExcelUrl/';
    public _getExcelCountUrl: string = '/Budget/CountYear';
    public _deleteExcelByIdUrl: string = '/Budget/DeleteCategoryById';
    public _formulaSheetData: wjcGridSheet.FlexSheet;
    public _staticData: Array<Array<Year>>;

    constructor(private http: Http) { }

   
    //Post
    saveExcel(formulaSheet: wjcGridSheet.FlexSheet): Observable<string> {
        let body = JSON.stringify(formulaSheet.saveToWorkbookOM());
        let headers = new Headers({ 'Content-Type': 'application/json' });
        let options = new RequestOptions({ headers: headers });

        return this.http.post(this._saveExcelUrl, body, options)
            .map(res => res["_body"])
            .catch(this.handleError);
    }
    saveExcel2(year: Array<Array<Year>>): Observable<string> {
        let body = JSON.stringify(year);
        let headers = new Headers({ 'Content-Type': 'application/json' });
        let options = new RequestOptions({ headers: headers });

        return this.http.post(this._saveExcel2Url, body, options)
            .map(res => res["_body"])
            .catch(this.handleError);
    }
    //Post
    getExcel(): Observable<string> {
        let body = "";
        let headers = new Headers({ 'Content-Type': 'application/json' });
        let options = new RequestOptions({ headers: headers });

        return this.http.post(this._getExcelUrl, body, options)
            .map(res => (res["_body"]))
            .catch(this.handleError);
    }
    getExcelCount(): Observable<Number> {
        var headers = new Headers();
        headers.append("If-Modified-Since", "Tue, 24 July 2017 00:00:00 GMT");
        var getExcelCount = this._getExcelCountUrl;
        return this.http.get(getExcelCount, { headers: headers })
            .map(response => <any>(<Response>response).json());
    }
    //Delete
    deleteExcelCat(id: Number[]): Observable<string> {
        let body = JSON.stringify(id);
        let headers = new Headers({ 'Content-Type': 'application/json' });
        let options = new RequestOptions({ headers: headers });
        return this.http.put(this._deleteExcelByIdUrl, body, options)
            .map(res => res["_body"])
            .catch(this.handleError);
    }
    private handleError(error: Response) {
        return Observable.throw(error.json().error || 'Opps!! Server error');
    }


}