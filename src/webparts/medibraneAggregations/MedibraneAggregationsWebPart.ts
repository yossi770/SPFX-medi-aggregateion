import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MedibraneAggregationsWebPart.module.scss';
import * as strings from 'MedibraneAggregationsWebPartStrings';


import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';




export interface IMedibraneAggregationsWebPartProps {
  description: string;
}

export default class MedibraneAggregationsWebPart extends BaseClientSideWebPart<IMedibraneAggregationsWebPartProps> {

  html:string;
  protected get isRenderAsync(): boolean {
    return true;
  }

  protected renderCompleted(): void {
    console.log("renderCompleted");
    this.domElement.innerHTML = `
      <div class="${ styles.medibraneAggregations }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div>
              ${this.html}
            </div>
          </div>
        </div>
      </div>`;
      //<h2>Loading Data 11</h2><h2>Loading Data</h2><h2>Loading Data</h2>
      //

    super.renderCompleted();

    //this.domElement.innerHTML = this.html;
    //this.domElement.innerHTML = '<h2>Loading Data</h2><h2>Loading Data</h2><h2>Loading Data</h2>'

    console.log("renderCompleted", this.domElement);
  }

  public renderError(){
    console.log("renderError", this);

  }

  public render(): void {
/*
    this.getListItems('Quotes').then((items)=>{});
    this.getListItems('Orders').then((items)=>{});
    this.getListItems('Projects').then((items)=>{});
*/

    this.domElement.innerHTML = '<h2>Loading Data</h2>'

    this.getListItems('Quotes');
    this.getListItems('Orders');
    this.getListItems('Projects');
    this.getListItems('Invoices');
    this.getListItems('Leads');
    this.getListItems('Expectations');


  }
//*************************globals****************************/
  ajaxCounter:number = 0;
  listsContainer:{} = {};
  today:Date = new Date();
  mm:number = (this.today.getMonth() + 1); //January is 0!


  public buildHtml(){
    console.log('buildHtml start');

    let monthlyQuotes = [];
    let monthlyInvoices = [];
    let nextMonthInvoices = [];
    let monthlyLeads = [];
    let lastMonthQuotes = [];
    let QuotesWaitingForResponse = [];
    let OrdersNotDelivered = [];
    let monthlyOrders = [];

//*********************created value for all functions****************************/
    let Created = (item , value:string) => {
      console.log('Created start');

      let createdFullVal = item[value];
      if(createdFullVal==null){
        return -1;
      }
      let month:number = createdFullVal[5]+createdFullVal[6];
      return month;

      return -1;//no month like -1
    };

//*****************start of weeks revenue function***************************/
let getWeek = (d) =>{
      let date:Date = new Date(d);
      let thisDay = date.getDate();
      console.log("today is ", thisDay);
      let thisMonth = date.getMonth();
      let thisYear = date.getFullYear();
      let startDayOfMonth = new Date(thisYear,thisMonth,1);

      console.log(startDayOfMonth.toString())

      console.log("the one's day ", startDayOfMonth.getDay()+1)


      let week = Math.ceil((thisDay+startDayOfMonth.getDay())/7);
      console.log("the returned week"+week);
      if(startDayOfMonth.getDay()+1 == 6 || startDayOfMonth.getDay()+1 == 7){
        /*if the first day is friday or saturday, the week starts in the next week*/
        console.log("first day of month is friday or saturday",week);
        return week-1;
      }
      return week;
    };

//*****************end of weeks revenue function***************************/


    //*****************leads count, and every level count***************************/
    let LeadsLevels = (arr:[], fName:string) => {
      console.log('Created LeadsLevels');

      let count = 0;
          let a = 0;
          let b = 0;
          let c = 0;
          let d = 0;
        for (let i = 0; i < arr.length; i++) {
          const item = arr[i];
          let level = item['Level'];
          let createdMonth = Created(item,'Created');
          if(createdMonth == this.mm){
            count++;
            if(level == 'a'){a++;}
            if(level == 'b'){b++;}
            if(level == 'c'){c++;}
            if(level == 'd'){d++;}
          }
        }
        return[count ,a ,b ,c ,d];
    };

    /***************************quotes from this month and the last one**********/
    let QuotesWon = (arr:[], month:number) => {
      console.log('QuotesWon');

      if(month == 0){
        month = 12;
      }
      let count:number = 0;
      let countSum:number = 0;
      let countWon:number = 0;
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        let created = Created(item,'Created');
        let QuotaAmount = item['Quota_x0020_amount'];
        if(created == month){
          count++;
          countSum += QuotaAmount;
          if(item['Quota_x0020_status'] == "Won"){
            countWon++;
          }
        }
      }
      let percent:number = countWon*100/count;
      if(count==0){percent = 0}
      return [countSum ,percent.toFixed(0)]

    }
      /********orders this month compared to expectations and projs this month*******/
    let invoicesCompared = (status:string) => {
      console.log('invoicesCompared start');

      let nextMonth =this.mm+1 ==13 ? 1:this.mm+1;
      let iArr = this.listsContainer['Invoices']
      let pArr = this.listsContainer['Projects']
      let eArr = this.listsContainer['Expectations']
      let monthly_Projects = 0;
      let week = [0,0,0,0,0,0];
      let monthly_Invoices = 0;
      let incomeExpectations = 0;
      if(status == '1'){
        for (let i = 0; i < iArr.length; i++) {
          const item = iArr[i];
          let createdMonth = Created(item,'Created');
          let status = item['Invoice_x0020_Status'];
          if(createdMonth == this.mm){
            if (status == 'An invoice was issued'){
              monthly_Invoices+=item['Order_x0020_Amount'];
            }
          }
        }
      }
      /**********changes in this loop - revenue seperate by weeks************/
      for (let i = 0; i < pArr.length; i++) {
        const item = pArr[i];
        let deliveryMonth = Created(item,'Delivery_x0020_Date');
        if(deliveryMonth == this.mm && status =='1' || deliveryMonth == nextMonth && status =='0'){
          if(item['Order_x0020_Amount']!=null){
            console.log(item['Order_x0020_Amount'])
            monthly_Projects+=item['Order_x0020_Amount'];
            /*start of treat the revenue seperate by weeks in loop*/
            let weekIndex = getWeek(item['Delivery_x0020_Date']);
            console.log(weekIndex-1);
            if(weekIndex>=0){
              week[weekIndex-1]+=item['Order_x0020_Amount'];
            }
            /*end of treat the revenue seperate by weeks in loop*/
            console.log('monthly projects' , monthly_Projects)
          }
        }
      }
      for (let i = 0; i < eArr.length; i++) {
        const item = eArr[i];
        let exMonth = Created(item,'Date1');
        if(exMonth == this.mm&& status =='1' || exMonth == nextMonth&& status =='0'){
          incomeExpectations = item['Monthly_x0020_income_x0020_forec'];
        }
      }
      console.log(monthly_Invoices+" "+incomeExpectations+" "+monthly_Projects)
      if(status == '1'){
        return [monthly_Invoices,incomeExpectations,monthly_Projects, week];/**param week- for weeks revenue. if remove this part, remove this param */
      }
      console.log(incomeExpectations+" "+monthly_Projects)
      return [monthly_Projects,incomeExpectations,week];/**param week- for weeks revenue. if remove this part, remove this param */
    };


          /************************************two parameters which are filtered by status***********************************/
    let filterByStatus = (arr:[], status:string) => {
      console.log('filterByStatus start');

      let returnVal = 0;
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        if(status == 'Waiting for customer response' && item['Quota_x0020_status'] == status &&item['Quota_x0020_amount']!=null){
          returnVal+=item['Quota_x0020_amount']
        }
        if(status == 'not finished' &&(item['Order_x0020_status'] == 'received'||item['Order_x0020_status'] == 'transferred to execution') &&item['Order_x0020_Amount']!=null){
          returnVal+=item['Order_x0020_Amount']
        }

      }
      console.log(returnVal)
      return[returnVal]
    }

    /************************************returns orders count and amount,compares to expectations***********************************/
    let OrdersAndExpectations = (arr1 , arr2) =>{
      console.log('OrdersAndExpectations start');

      let count:number = 0;
      let countSum:number = 0;
      let expectedOrders:number = 0;
      for (let i = 0; i < arr1.length; i++) {
        const item = arr1[i];
        let created = Created(item,'Created');
        let orderAmount = item['Order_x0020_Amount'];
        if(created == this.mm){
          count++;
          countSum += orderAmount;
        }
      }
      for(let i = 0; i < arr2.length; i++) {
        const item = arr2[i];
        let created = Created(item,'Created');
        if(created == this.mm){
          expectedOrders += item['Expect_x0020_monthly_x0020_order']
        }
      }
      return [count ,expectedOrders, countSum]
    }



    //*********************return array of amount and count**********************//
    monthlyQuotes = QuotesWon(this.listsContainer['Quotes'], this.mm );
    lastMonthQuotes = QuotesWon(this.listsContainer['Quotes'], this.mm-1);
    monthlyInvoices = invoicesCompared('1');
    nextMonthInvoices = invoicesCompared('0');
    monthlyLeads = LeadsLevels(this.listsContainer['Leads'], 'Level')
    QuotesWaitingForResponse = filterByStatus(this.listsContainer['Quotes'] , 'Waiting for customer response')
    OrdersNotDelivered = filterByStatus(this.listsContainer['Orders'] , 'not finished')
    monthlyOrders = OrdersAndExpectations(this.listsContainer['Orders'] , this.listsContainer['Expectations'])

    console.log('setting this.domElement.innerHTML');

    //this.domElement.innerHTML = `
    this.html = `
      <div>
        <div>
          <div class="${styles.flexCenterText}"}>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Monthly Leads </label>
              </div>
              Leads amount : ${monthlyLeads[0]}</br>
              level a :  ${monthlyLeads[1]}</br>
              level b :  ${monthlyLeads[2]}</br>
              level c :  ${monthlyLeads[3]}</br>
              level d :  ${monthlyLeads[4]}</br>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Monthly Quotes </label></br>
              </div>
              quotes amount : ${monthlyQuotes[0]}</br>
              quotes won : ${monthlyQuotes[1]}%</br>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Quotes last month</label></br>
              </div>
              quotes amount : ${lastMonthQuotes[0]}</br>
              won quotes : ${lastMonthQuotes[1]}%</br>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Revenues this month</label> </br>
              </div>
              Invoices amount : ${monthlyInvoices[0]}</br>
              Projects amount : ${monthlyInvoices[1]}</br>
              Revenue expected :  ${monthlyInvoices[2]}</br>
              <table>
                <tr><td>week1</td><td>week2</td><td>week3</td><td>week4</td><td>week5</td></tr>
                <tr>
                  <td>${monthlyInvoices[3][0]}</td>
                  <td>${monthlyInvoices[3][1]}</td>
                  <td>${monthlyInvoices[3][2]}</td>
                  <td>${monthlyInvoices[3][3]}</td>
                  <td>${monthlyInvoices[3][4]}</td>
                </tr>
              </table>
            </div>

          </div>
        </div>

        <div>
          <div class="${styles.flexCenterText}"}>


            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Revenues next month</label></br>
              </div>
              Revenue expected : ${nextMonthInvoices[0]}</br>
              expected income  : ${nextMonthInvoices[1]}</br>
              <table>
                <tr>
                  <td>week1</td><td>week2</td><td>week3</td><td>week4</td><td>week5</td>
                </tr>
                <tr>
                  <td>${nextMonthInvoices[2][0]}</td>
                  <td>${nextMonthInvoices[2][1]}</td>
                  <td>${nextMonthInvoices[2][2]}</td>
                  <td>${nextMonthInvoices[2][3]}</td>
                  <td>${nextMonthInvoices[2][4]}</td>
                </tr>
              </table>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Quotations waiting</label></br>
              </div>
              quotes amount : ${QuotesWaitingForResponse[0]}</br>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>Not finished orders</label></br>
              </div>
              orders amount : ${OrdersNotDelivered[0]}</br>
            </div>

            <div class="${ styles.SumsDiv }">
              <div class = "${ styles.labelDiv }">
                <label>monthly Orders</label></br>
              </div>
              number of orders  : ${monthlyOrders[0]}</br>
              expected number  : ${monthlyOrders[1]}</br>
              orders amount  : ${monthlyOrders[2]}</br>
            </div>

          </div>
        </div>

      </div>`;

    console.log(this.domElement);
    this.renderCompleted();

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }


//**************************** returns the full lists *************************/
    public getListItems(listname:string): void {

      console.log('asking list items for', listname);
      this.ajaxCounter++;

      this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
        `/_api/web/lists/GetByTitle('${listname}')/Items`, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
                response.json().then((data)=> {

                    console.log('list items for', listname, data);
                    this.ajaxCounter--;
                    this.listsContainer[listname] = data.value;
                    if (this.ajaxCounter == 0) {
                      this.buildHtml();
                    }

                });
            });
      }

}



//  public getListItems(listname:string): Promise<{}[]> {
//    console.log('asking list items for', listname);
//    this.ajaxCounter++;
//
//    return this.context.spHttpClient.get(
//      this.context.pageContext.web.absoluteUrl +
//      `/_api/web/lists/GetByTitle('${listname}')/Items`, SPHttpClient.configurations.v1)
//          .then((response: SPHttpClientResponse) => {
//              let items = response.json()['value'];
//              console.log('list items for', listname, items);
//              this.ajaxCounter--;
//              if (this.ajaxCounter == 0) {
//
//              }
//              return items;
//          });
//    }

