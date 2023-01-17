/* tslint:disable */
import { Log, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

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

  html: string;
  protected get isRenderAsync(): boolean {
    return true;
  }

  protected renderCompleted(): void {
    // console.log("renderCompleted");
    this.domElement.innerHTML = `
      <div class="${styles.medibraneAggregations}">
        <div class="${styles.container}">
          <div class="${styles.row}">
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

    // console.log("renderCompleted", this.domElement);
  }

  public renderError() {
    // console.log("renderError", this);

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
    this.getListItems('Projects');//
    //this.getListItems('Application Development Lab (ADL)', 'Projects');
    //this.getListItems('Advanced Development', 'Projects');
    //this.getListItems('Manufacturing', 'Projects');

    this.getListItems('Invoices');
    this.getListItems('Leads');
    this.getListItems('Expectations');

    this.getListItems('Shipments');
    this.getListItems('Supplier  Purchases');
    this.getListItems('Employees-Projects-Hours');

  }

  //*************************globals****************************/
  ajaxCounter: number = 0;
  listsContainer: {} = {};
  today: Date = new Date();
  mm: number = (this.today.getMonth() + 1); //January is 0!


  public buildHtml() {
    // console.log('buildHtml start');

    let monthlyQuotes = [];
    let monthlyInvoices = [];
    let nextMonthInvoices = [];
    let monthlyLeads = [];
    let lastMonthQuotes = [];
    let QuotesWaitingForResponse = [];
    let OrdersNotDelivered = [];
    let monthlyOrders = [];

    //*********************created value for all functions****************************/
    let Created = (item, value: string) => {
      // console.log('Created start');

      let createdFullVal: Date = new Date(item[value]);
      if (createdFullVal == null) {
        return -1;
      }
      let month: number = createdFullVal.getMonth() + 1;
      // console.log('the created is: '+month);
      return month;

      return -1;//no month like -1
    };

    //*****************start of weeks revenue function***************************/
    let getWeek = (d) => {
      let date: Date = new Date(d);
      let thisDay = date.getDate();
      // console.log("getWeek :: today is ", thisDay);
      let thisMonth = date.getMonth();
      let thisYear = date.getFullYear();
      let startDayOfMonth = new Date(thisYear, thisMonth, 1);

      // console.log(startDayOfMonth.toString())

      // console.log("the one's day ", startDayOfMonth.getDay()+1)


      let week = Math.ceil((thisDay + startDayOfMonth.getDay()) / 7);
      // console.log("the returned week"+week);
      if (startDayOfMonth.getDay() + 1 == 6 || startDayOfMonth.getDay() + 1 == 7) {
        /*if the first day is friday or saturday, the week starts in the next week*/
        // console.log("first day of month is friday or saturday",week);
        return week - 1;
      }
      return week;
    };

    //*****************end of weeks revenue function***************************/


    //*****************leads count, and every level count***************************/
    let LeadsLevels = (arr: [], fName: string) => {
      // console.log('Created LeadsLevels');

      let count = 0;
      let a = 0;
      let b = 0;
      let c = 0;
      let d = 0;
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        let level = item['Level'];
        // let createdMonth = Created(item,'Created');
        if (new Date(item['Created']).getMonth() + 1 === new Date().getMonth() + 1 && new Date(item['Created']).getFullYear() === new Date().getFullYear()) {
          count++;
          if (level == 'a') { a++; }
          if (level == 'b') { b++; }
          if (level == 'c') { c++; }
          if (level == 'd') { d++; }
        }
      }
      return [count, a, b, c, d];
    };

    /***************************quotes from this month and the last one**********/
    let QuotesWon = (arr: [], action: string) => {
      // console.log('QuotesWon');
      let month = new Date().getMonth() + 1;
      let year = new Date().getFullYear();
      if (action === 'lastMonthQuotes') {
        const LessMonth = addMonths(new Date(), -1);
        month = new Date(LessMonth).getMonth() + 1;
        year = new Date(LessMonth).getFullYear();
      }

      let expectArr: [] = this.listsContainer['Expectations'];
      //console.log("quotes exp ", expectArr)*******
      let qoutesEx: number;
      let count: number = 0;
      let countSum: number = 0;
      let countWon: number = 0;
      let percent: number = 0;
      let seperateByTypes = [0, 0, 0]
      let wonByTypes = [0, 0, 0];
      let wonPercent = [0, 0, 0]
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        //let created = Created(item,'Created');
        // let created = Created(item,'project_x0020_sending_x0020_date');

        let QuotaAmount = item['Quota_x0020_amount'];
        if (new Date(item['project_x0020_sending_x0020_date']).getMonth() + 1 === month && new Date(item['project_x0020_sending_x0020_date']).getFullYear() === year) {
          // if(created == month){
          count++;
          switch (item['Lead_x0020_type']) {
            case 'II':
              seperateByTypes[0]++;
              if (item['Quota_x0020_status'] == "Won") {
                wonByTypes[0]++;
              }
              break;
            case 'III':
              seperateByTypes[1]++;
              if (item['Quota_x0020_status'] == "Won") {
                wonByTypes[1]++;
              }
              break;
            case 'IV':
              seperateByTypes[2]++;
              if (item['Quota_x0020_status'] == "Won") {
                wonByTypes[2]++;
              }
              break;
          }

          countSum += QuotaAmount;
          if (item['Quota_x0020_status'] == "Won") {
            countWon++;
          }
        }
      }

      // console.log("the seperate arr",seperateByTypes[0],seperateByTypes[1],seperateByTypes[2])

      for (let i = 0; i < 3; i++) {
        if (seperateByTypes[i] != 0) {
          wonPercent[i] = parseInt(((wonByTypes[i] * 100) / seperateByTypes[i]).toFixed(0));
        }
        else {
          wonPercent[i] = 0;
        }
        // console.log("wonPercent",i,wonPercent[i]);
      }

      if (count != 0) {
        percent = countWon * 100 / count;
      }

      for (let i = 0; i < expectArr.length; i++) {
        const item = expectArr[i];
        let exMonth = Created(item, 'Date1');
        if (exMonth == month) {
          qoutesEx = item['QuotesBudges'];
        }
      }
      // console.log("the quotes amount for month ", month, " is ", qoutesEx);



      return [countSum.toFixed(0), percent.toFixed(0), wonPercent, qoutesEx];

    }

    // Backlog fun
    function addMonths(date, months) {
      var d = date.getDate();
      date.setMonth(date.getMonth() + +months);
      if (date.getDate() != d) {
        date.setDate(0);
      }
      return date;
    }


    let EmployeesHours = this.listsContainer['Employees-Projects-Hours']
    // console.log(EmployeesHours, 'EmployeesHours');
    let Shipments = this.listsContainer['Shipments']
    let Supplier = this.listsContainer['Supplier  Purchases']
    let ProArr = this.listsContainer['Projects']
    let projectsThisMonth = ProArr.filter(i => new Date(i.Delivery_x0020_Date).getMonth() === new Date().getMonth() && new Date(i.Delivery_x0020_Date).getFullYear() === new Date().getFullYear())
    // console.log(projectsThisMonth.length, 'projectsThis', projectsThisMonth[0]);
    let RevenueFromLabor = 0;
    let ShipmentsCost = 0;
    let SupplierCost = 0;
    let rawmaterialcost = 0;
    let IncomeProject = 0;
    let EmployeesHoursCost = 0;
    let araayOfcont = [];
    let avgByTeamType = [];
    let IncomeProjectByTeamType = [];
    let contvrr = 0
    for (let index = 0; index < projectsThisMonth.length; index++) {
      let projectShipmentsCost = 0;
      let projectSupplierCost = 0;
      let ProjectEmployeesHoursCost = 0;
      const element = projectsThisMonth[index];
      if (element["Order_x0020_Amount"]) {
        contvrr++
        IncomeProject += element["Order_x0020_Amount"];
        let ex = IncomeProjectByTeamType.filter(i => i.TeamType === element.ProjectTeamType)
        if (ex.length > 0) {
          ex[0].value += element["Order_x0020_Amount"]
        }
        else {
          IncomeProjectByTeamType.push({ TeamType: element.ProjectTeamType, value: element["Order_x0020_Amount"] })
        }
        let projectShipments = Shipments.filter(i => i.Tag_x0020_NumberId === element.ID);
        let projectSupplier = Supplier.filter(i => i.Tag_x0020_NumberId === element.ID);
        let projectEmployeesHours = EmployeesHours.filter(i => i.Project_x0020_lookupId === element.ID);
        if (EmployeesHours.length > 0) {
          for (let index = 0; index < projectEmployeesHours.length; index++) {
            const elementEmployees = projectEmployeesHours[index];
            if (elementEmployees.TotalMoney) {
              EmployeesHoursCost += elementEmployees.TotalMoney;
              ProjectEmployeesHoursCost += elementEmployees.TotalMoney;
            }
          }
        }
        if (projectShipments.length > 0) {
          for (let index = 0; index < projectShipments.length; index++) {
            const elementShipments = projectShipments[index];
            if (elementShipments.TotalPrice) {
              ShipmentsCost += elementShipments.TotalPrice;
              projectShipmentsCost += elementShipments.TotalPrice;
            }
          }
        }
        if (projectSupplier.length > 0) {
          for (let index = 0; index < projectSupplier.length; index++) {
            const elementSupplier = projectSupplier[index];
            if (elementSupplier.TotalPrice) {
              SupplierCost += elementSupplier.TotalPrice;
              projectSupplierCost += elementSupplier.TotalPrice;
            }
          }
        }
        if (element.rawmaterialcost) {
          rawmaterialcost += element.rawmaterialcost;
        }

        let externalIncomeProject = (projectShipmentsCost + projectSupplierCost + element.rawmaterialcost) * 1.2;
        let RevenueFromLaborProject = element["Order_x0020_Amount"] - externalIncomeProject - ProjectEmployeesHoursCost;
        let percentagesRevenueFromLaborProject = RevenueFromLaborProject / element["Order_x0020_Amount"] * 100
        araayOfcont.push({ ProjectTeamType: element.ProjectTeamType, IncomeOneProject: element["Order_x0020_Amount"], contributionMargin: percentagesRevenueFromLaborProject })
      }
    }
    // console.log(contvrr, 'contvrr');

    // console.log(araayOfcont);
    let avg = 0;
    // Application Development Lab (ADL)
    // Manufacturing
    // Quality
    // Engineering
    // Raw Materials
    // *********************
    let projectsYear2021 = ProArr.filter(i => new Date(i.Delivery_x0020_Date).getFullYear() === 2021)
    let projectsYear2022 = ProArr.filter(i => new Date(i.Delivery_x0020_Date).getFullYear() === 2022)
    
    let SupplierCost2021 = 0;
    let SupplierCost2022 = 0;
    let rawmaterialCost2021 = 0;
    let rawmaterialCost2022 = 0;
    projectsYear(2021, projectsYear2021,SupplierCost2021, rawmaterialCost2021)
    projectsYear(2022, projectsYear2022, SupplierCost2022,rawmaterialCost2022 )
    function projectsYear (year, list, Suppliercost, rawmaterialcostforyear) {
      for (let index = 0; index < list.length; index++) {
        const element = list[index];
        // if (element["Order_x0020_Amount"]) {
          let projectSupplier = Supplier.filter(i => i.Tag_x0020_NumberId === element.ID);
          if (projectSupplier.length > 0) {
            for (let index = 0; index < projectSupplier.length; index++) {
              const elementSupplier = projectSupplier[index];
              if (elementSupplier.TotalPrice) {
                Suppliercost += elementSupplier.TotalPrice;
              }
            }
          }
          if (element.rawmaterialcost) {
            rawmaterialcostforyear += element.rawmaterialcost;
            
          }
        // }
      }
      console.log('year:',year , 'Suppliercost:',Number(Suppliercost.toFixed(0)).toLocaleString('en') , 'rawmaterialCost:', Number(rawmaterialcostforyear.toFixed(0)).toLocaleString('en')  );
    }

    // *****************
    for (let index = 0; index < araayOfcont.length; index++) {
      const element = araayOfcont[index];
      const Weight = element.IncomeOneProject / IncomeProject;
      const ValueByWeight = Weight * element.contributionMargin;
      avg += ValueByWeight;
      const WeightByTeamType = element.IncomeOneProject / IncomeProjectByTeamType.filter(i => i.TeamType === element.ProjectTeamType)[0].value;
      const ValueByWeightByTeamType = WeightByTeamType * element.contributionMargin;
      // IncomeProjectByTeamType
      let e = avgByTeamType.filter(i => i.TeamType === element.ProjectTeamType)
      if (e.length > 0) {
        e[0].value += ValueByWeightByTeamType
      }
      else {
        avgByTeamType.push({ TeamType: element.ProjectTeamType, value: ValueByWeightByTeamType })
      }
      // element.ProjectTeamType
      // console.log('IncomeProjectByTeamType:',IncomeProjectByTeamType,'ProjectTeamType:', element.ProjectTeamType, 'Income:', element.IncomeOneProject, 'Weight:', Weight, 'contributionMargin:', element.contributionMargin, 'ValueByWeight:', ValueByWeight);

    }
    // console.log(avg, 'avg');
    // console.log(avgByTeamType, 'avgByTeamType');
    let avgByTeamTypeHtml = '';
    avgByTeamType.forEach(element =>
      avgByTeamTypeHtml += `<span> ${element.TeamType === 'Application Development Lab (ADL)' ? 'ADL' : element.TeamType}: ${element.value.toFixed(0)}%</span></br>`
    )
    let Expectations = this.listsContainer['Expectations'].filter(i => new Date(i.Date1).getMonth() === new Date().getMonth() && new Date(i.Date1).getFullYear() === new Date().getFullYear());
    // console.log(Expectations, 'Expectations');
    let numWorkingDays = 0;
    let NumEmployees = 0;
    if (Expectations.length > 0) {
      if (Expectations[0].numWorkingDays) { numWorkingDays = Expectations[0].numWorkingDays; }
      if (Expectations[0].NumEmployees) { NumEmployees = Expectations[0].NumEmployees; }
    }
    // numWorkingDays
    // let f = Expectations.filter(i => new Date(i.Date1).getMonth() === new Date().getMonth() && new Date(i.Date1).getFullYear() === new Date().getFullYear())
    // console.log(f, 'Expectations f');
    console.log('EmployeesHoursCost:', EmployeesHoursCost, 'ShipmentsCost:', ShipmentsCost, 'SupplierCost:', SupplierCost, 'rawmaterialcost:', rawmaterialcost, 'IncomeProject:', IncomeProject);
    // SupplierCost + rawmaterialcost  2022//
    const RevenueFromExternalResources = ( SupplierCost + rawmaterialcost) * 1.2;
    console.log('RevenueFromExternalResources', RevenueFromExternalResources, SupplierCost, rawmaterialcost);
    
    const externalIncomeCost = (ShipmentsCost + SupplierCost + rawmaterialcost) * 1.2;
    // RevenueFromLabor += (IncomeProject - externalIncomeCost - EmployeesHoursCost);
    RevenueFromLabor += (IncomeProject - externalIncomeCost);
    // console.log(RevenueFromLabor, 'RevenueFromLabor');
    const RevenueFromLaborBy = RevenueFromLabor / (numWorkingDays * NumEmployees) //by employees by working days
    console.log(RevenueFromLaborBy, 'RevenueFromLaborBy');
    // console.log(RevenueFromLabor, numWorkingDays, NumEmployees, 'RevenueFromLaborBy num');// לבדוק איך לחלק את זה

    let projects = ProArr.filter(i => i.Project_x0020_Status !== "ended" && i.Project_x0020_Status !== "Sent to customer")
    let nextOneMonthRevenue = 0;
    let nextTwoMonthRevenue = 0;
    for (let index = 0; index < projects.length; index++) {
      const element = projects[index];
      const addOnemonths = addMonths(new Date(), 1);
      const OneM = new Date(addOnemonths).getMonth() + 1;
      const OneY = new Date(addOnemonths).getFullYear();
      const addTwomonths = addMonths(new Date(), 2);
      const TwoM = new Date(addTwomonths).getMonth() + 1;
      const TwoY = new Date(addTwomonths).getFullYear();
      if (new Date(element.Delivery_x0020_Date).getMonth() + 1 === OneM && new Date(element.Delivery_x0020_Date).getFullYear() === OneY) {
        nextOneMonthRevenue += element.Order_x0020_Amount
      }
      if (new Date(element.Delivery_x0020_Date).getMonth() + 1 === TwoM && new Date(element.Delivery_x0020_Date).getFullYear() === TwoY) {
        nextTwoMonthRevenue += element.Order_x0020_Amount
      }
    }
    let OrdersSum: number = 0;
    let ProjectsSum: number = 0;
    let Orders = this.listsContainer['Orders'].filter(i => true === (i["Order_x0020_status"] != "sent to customer" && i["Order_x0020_status"] != "ended"))
    let Projects = this.listsContainer['Projects'].filter(i => true === (i["Project_x0020_Status"] != "Summary" && i["Project_x0020_Status"] != "Sent to customer" && i["Project_x0020_Status"] != "ended"))
    for (let i = 0; i < Projects.length; i++) {
      const item = Projects[i];
      if (item["Order_x0020_Amount"]) {
        // console.log(item["Order_x0020_Amount"]);
        ProjectsSum = ProjectsSum + item["Order_x0020_Amount"]
      }
    }
    for (let i = 0; i < Orders.length; i++) {
      const item = Orders[i];
      if (item["lefttopay"]) {
        // console.log(item["lefttopay"]);
        OrdersSum = OrdersSum + item["lefttopay"]
      }
    }

    /********orders this month compared to expectations and projs this month*******/
    let invoicesCompared = (status: string) => {
      console.log('invoicesCompared start');
      //debugger

      let nextMonth = this.mm + 1 == 13 ? 1 : this.mm + 1;
      let iArr = this.listsContainer['Invoices']
      let pArr = this.listsContainer['Projects']
      let eArr = this.listsContainer['Expectations']
      let monthly_Projects = 0;
      let projectWeek = [0, 0, 0, 0, 0];
      let invoiceWeek = [0, 0, 0, 0, 0];
      let PerformedInPercent = [0, 0, 0, 0, 0]
      let monthly_Invoices = 0;
      let incomeExpectations = 0;

      if (status == '1') {
        // console.log(iArr[1], iArr.length);

        // for (let index = 0; index < iArr.length; index++) {
        //   const element = iArr[index];
        //   let status = element['Invoice_x0020_Status'];
        //    if ( new Date(element.Created).getMonth()+1 === new Date().getMonth()+1 && new Date(element.Created).getFullYear() === new Date().getFullYear()){
        //        if (element.InvoiceAmount){
        //     if (status == 'An invoice was issued'){
        //         monthly_Invoices += element.InvoiceAmount           
        //        } }
        //    }
        //  }

        for (let i = 0; i < iArr.length; i++) {
          // const item = iArr[i];
          const element = iArr[i];
          // let createdMonth = Created(item,'Created');
          let status = element['Invoice_x0020_Status'];
          if (new Date(element.PracticalDate).getMonth() + 1 === new Date().getMonth() + 1 && new Date(element.PracticalDate).getFullYear() === new Date().getFullYear()) {
            if (status == 'An invoice was issued') {
              monthly_Invoices += element['InvoiceAmount'];
            }
            let weekIndex = getWeek(element['PracticalDate']);
            // console.log(weekIndex-1);
            if (weekIndex >= 0) {
              invoiceWeek[weekIndex - 1] += element['InvoiceAmount'];
            }
          }
        }
        // console.log('invoiceweek**************'+invoiceWeek[1])
      }
      /**********changes in this loop - revenue seperate by weeks************/
      // console.log('invoicesCompared for (let i = 0; i < pArr.length; i++)', pArr);
      for (let i = 0; i < pArr.length; i++) {
        const item = pArr[i];
        const addOnemonths = addMonths(new Date(), 1);
        const OMi = new Date(addOnemonths).getMonth() + 1;
        const OYi = new Date(addOnemonths).getFullYear();
        // let x:Date = new Date(item['Delivery_x0020_Date'])
        let delivery = item.Delivery_x0020_Date;
        // console.log('the x', deliveryMonth)


        // console.log('invoicesCompared if deliveryMonth ... ', ((deliveryMonth == this.mm && status =='1') || (deliveryMonth == nextMonth && status =='0')));
        if (new Date(delivery).getMonth() + 1 === new Date().getMonth() + 1 && new Date(delivery).getFullYear()
          === new Date().getFullYear() && status === '1' || new Date(delivery).getMonth() + 1 === OMi && new Date(delivery).getFullYear() === OYi && status === '0') {
          // if((deliveryMonth == this.mm && status =='1') || (deliveryMonth == nextMonth && status =='0')){

          // console.log('invoicesCompared if Order_x0020_Amount ... ', (item['Order_x0020_Amount']!=null));
          if (item['Order_x0020_Amount'] != null) {
            // console.log(item['Order_x0020_Amount'])
            monthly_Projects += item['Order_x0020_Amount'];
            /*start of treat the revenue seperate by weeks in loop*/
            let weekIndex = getWeek(item['Delivery_x0020_Date']);
            // console.log(weekIndex-1);
            if (weekIndex >= 0) {
              projectWeek[weekIndex - 1] += item['Order_x0020_Amount'];
            }
            /*end of treat the revenue seperate by weeks in loop*/
            if (status == '1') {
              // console.log('monthly projects, my month is' , monthly_Projects, item['Delivery_x0020_Date'],item['Order_x0020_Amount']);
            }
          }
        }
      }

      // console.log('monthly project amount ', monthly_Projects);

      for (let i = 0; i < 5; i++) {
        projectWeek[i].toFixed(0);
        if (projectWeek[i] == 0) {
          PerformedInPercent[i] = 0;
        }
        else {
          PerformedInPercent[i] = ((invoiceWeek[i] * 100) / projectWeek[i]);
          PerformedInPercent[i] = Number(PerformedInPercent[i].toFixed(0));
        }
        // console.log("------------the preformed by weeks",PerformedInPercent[i]);
      }

      for (let i = 0; i < eArr.length; i++) {
        const item = eArr[i];
        // let exMonth = Created(item,'Date1');
        const addOnemonths = addMonths(new Date(), 1);
        const OM = new Date(addOnemonths).getMonth() + 1;
        const OY = new Date(addOnemonths).getFullYear();
        if (new Date(item['Date1']).getMonth() + 1 === new Date().getMonth() + 1 && new Date(item['Date1']).getFullYear()
          === new Date().getFullYear() && status == '1' || new Date(item['Date1']).getMonth() + 1 === OM && new Date(item['Date1']).getFullYear() === OY && status == '0') {
          incomeExpectations = item['Monthly_x0020_income_x0020_forec'];
        }
      }
      // console.log(monthly_Invoices+" "+incomeExpectations+" "+monthly_Projects)
      if (status == '1') {
        return [monthly_Invoices.toFixed(0), incomeExpectations.toFixed(0), monthly_Projects.toFixed(0), projectWeek, PerformedInPercent];/**param week- for weeks revenue. if remove this part, remove this param */
      }
      // console.log(incomeExpectations+" "+monthly_Projects)
      return [monthly_Projects.toFixed(0), incomeExpectations.toFixed(0), projectWeek];/**param week- for weeks revenue. if remove this part, remove this param */
    };


    /************************************two parameters which are filtered by status***********************************/
    let filterByStatus = (arr: [], status: string) => {
      // console.log('filterByStatus start');/**TODO column name = Activity_x0020_type */
      let countByActivityTypes = [0, 0, 0, 0];

      let returnVal = 0;
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        if (status == 'Waiting for customer response' && item['Quota_x0020_status'] == status && item['Quota_x0020_amount'] != null) {
          returnVal += item['Quota_x0020_amount']
        }
        if (status == 'not finished' && item['Order_x0020_status'] != 'ended' && item['Order_x0020_status'] != 'Sent to customer' && item['Order_x0020_Amount'] != null) {

          returnVal += item['Order_x0020_Amount']
          switch (item['Activity_x0020_type']) {
            case 'Prototyping':
              countByActivityTypes[0] += item['Order_x0020_Amount'];
              break;
            case 'Mprep':
              countByActivityTypes[1] += item['Order_x0020_Amount'];
              break;
            case 'Clinical Builds':
              countByActivityTypes[2] += item['Order_x0020_Amount'];
              break;
            case 'Manufacturing':
              countByActivityTypes[3] += item['Order_x0020_Amount'];
              break;
          }
        }

      }

      if (status == 'not finished') {
        for (let t = 0; t < countByActivityTypes.length; t++) {
          countByActivityTypes[t] = parseInt(countByActivityTypes[t].toFixed(0));
          // console.log("countByActivityTypes",t,countByActivityTypes[t]);
        }
      }
      // console.log("returnVal",returnVal)
      if (status == 'not finished') {
        return [returnVal.toFixed(0), countByActivityTypes];
      }
      return [returnVal.toFixed(0)];
    }

    /************************************returns orders count and amount,compares to expectations***********************************/
    let OrdersAndExpectations = (arr1, arr2) => {
      // console.log('OrdersAndExpectations start');

      let count: number = 0;
      let countSum: number = 0;
      let expectedOrders: number = 0;
      for (let i = 0; i < arr1.length; i++) {
        const item = arr1[i];
        // let created = Created(item,'Created');
        let orderAmount = item['Order_x0020_Amount'];
        if (new Date(item['Created']).getMonth() + 1 === new Date().getMonth() + 1 && new Date(item['Created']).getFullYear() === new Date().getFullYear()) {
          // if(created == this.mm){
          count++;
          countSum += orderAmount;
        }
      }
      for (let i = 0; i < arr2.length; i++) {
        const item = arr2[i];
        // let created = Created(item,'Date1');
        // if(created == this.mm){
        if (new Date(item['Date1']).getMonth() + 1 === new Date().getMonth() + 1 && new Date(item['Date1']).getFullYear() === new Date().getFullYear()) {
          expectedOrders += item['Expect_x0020_monthly_x0020_order']
        }
      }
      return [count, expectedOrders, countSum.toFixed(0)]
    }



    //*********************return array of amount and count**********************//
    monthlyQuotes = QuotesWon(this.listsContainer['Quotes'], 'monthlyQuotes');
    lastMonthQuotes = QuotesWon(this.listsContainer['Quotes'], 'lastMonthQuotes');
    monthlyInvoices = invoicesCompared('1');
    nextMonthInvoices = invoicesCompared('0');
    monthlyLeads = LeadsLevels(this.listsContainer['Leads'], 'Level')
    QuotesWaitingForResponse = filterByStatus(this.listsContainer['Quotes'], 'Waiting for customer response')
    OrdersNotDelivered = filterByStatus(this.listsContainer['Orders'], 'not finished')
    monthlyOrders = OrdersAndExpectations(this.listsContainer['Orders'], this.listsContainer['Expectations'])
    const monthNames = ["January", "February", "March", "April", "May", "June",
      "July", "August", "September", "October", "November", "December"
    ];
    const Onemonths = addMonths(new Date(), 1);
    const One = new Date(Onemonths).getMonth();
    const Twomonths = addMonths(new Date(), 2);
    const Two = new Date(Twomonths).getMonth();
    // console.log('setting this.domElement.innerHTML');
    //this.domElement.innerHTML = `
    this.html = `
      <div>
        <div>
          <div class="${styles.flexCenterText}"}>

            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Monthly leads </label>
              </div>
              Leads amount : ${Number(monthlyLeads[0]).toLocaleString('en')}</br>
              Level a :  ${Number(monthlyLeads[1]).toLocaleString('en')}</br>
              Level b :  ${Number(monthlyLeads[2]).toLocaleString('en')}</br>
              Level c :  ${Number(monthlyLeads[3]).toLocaleString('en')}</br>
              Level d :  ${Number(monthlyLeads[4]).toLocaleString('en')}</br>
            </div>

            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Monthly quotes </label></br>
              </div>
              Quotes amount : ${Number(monthlyQuotes[0]).toLocaleString('en')}</br>
              Quotes won : ${monthlyQuotes[1]}%</br>
              Quotes budget : ${monthlyQuotes[3].toLocaleString('en')}</br>
              </br>
              Quotes awaiting response : ${Number(QuotesWaitingForResponse[0]).toLocaleString('en')}</br>
            </div>

            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Quotes last month</label></br>
              </div>
              Quotes amount : ${Number(lastMonthQuotes[0]).toLocaleString('en')}</br>
              Won quotes : ${lastMonthQuotes[1]}%</br>
              <table>
              <tr>
                <td>Type II</td>
                <td>Type III</td>
                <td>Type IV</td>
              </tr>
              <tr>
                <td> ${lastMonthQuotes[2][0]}%</td>
                <td> ${lastMonthQuotes[2][1]}%</td>
                <td> ${lastMonthQuotes[2][2]}%</td>
              </tr>
            </table>

            </div>

            <div class="${styles.SumsDiv}" style="line-height: 18px;">
              <div class = "${styles.labelDiv}">
                <label>Revenues this month</label> </br>
              </div>
              Invoices amount : ${Number(monthlyInvoices[0]).toLocaleString('en')}</br>
              Revenue budget : ${Number(monthlyInvoices[1]).toLocaleString('en')}</br>
              Revenue expected :  ${Number(monthlyInvoices[2]).toLocaleString('en')}</br>
              <table>
                <tr><td>Week1</td><td>Week2</td><td>Week3</td><td>Week4</td><td>Week5</td></tr>
                <tr>
                  <td>${Number(monthlyInvoices[3][0]).toLocaleString('en')}</td>
                  <td>${Number(monthlyInvoices[3][1]).toLocaleString('en')}</td>
                  <td>${Number(monthlyInvoices[3][2]).toLocaleString('en')}</td>
                  <td>${Number(monthlyInvoices[3][3]).toLocaleString('en')}</td>
                  <td>${Number(monthlyInvoices[3][4]).toLocaleString('en')}</td>
                </tr>
                <tr>
                  <td>${monthlyInvoices[4][0]}%</td>
                  <td>${monthlyInvoices[4][1]}%</td>
                  <td>${monthlyInvoices[4][2]}%</td>
                  <td>${monthlyInvoices[4][3]}%</td>
                  <td>${monthlyInvoices[4][4]}%</td>
                </tr>

              </table>
            </div>

          </div>
        </div>

        <div>
          <div class="${styles.flexCenterText}"}>


            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Revenues next months</label></br>
              </div>
              ${monthNames[One]} revenue expected  : ${nextOneMonthRevenue.toLocaleString('en')}</br>
              ${monthNames[Two]} revenue expected  : ${nextTwoMonthRevenue.toLocaleString('en')}</br>
            </div>

            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Monthly labor rev</label></br>
              </div>
              Monthly labor from rev: ${Number(RevenueFromLabor.toFixed(0)).toLocaleString('en')}</br>
              Revenue/employees/working days: ${Number(RevenueFromLaborBy.toFixed(0)).toLocaleString('en')}</br>
              Working: ${numWorkingDays}</br>
              Employees: ${NumEmployees}</br>
              Revenue from external resources: ${Number(RevenueFromExternalResources.toFixed(0)).toLocaleString('en')}
            </div>

            <div class="${styles.SumsDiv}">
              <div class = "${styles.labelDiv}">
                <label>Avg CM</label></br>
              </div>
              Total : ${avg.toFixed(0)}%</br>
              ${avgByTeamTypeHtml}
            </div>

            <div class="${styles.SumsDiv}" style="line-height: 18px;">
              <div class = "${styles.labelDiv}">
                <label>Monthly orders</label></br>
              </div>
              Number of orders  : ${monthlyOrders[0].toLocaleString('en')}</br>
              Po's budget  : ${monthlyOrders[1].toLocaleString('en')}</br>
              Orders amount  : ${monthlyOrders[2].toLocaleString('en')}</br>
              <div class = "${styles.labelDiv}" style="margin-top: 0; height: 30px;">
                    <label>Backlog</label></br>
              </div>
              Open order amount  : ${Number(OrdersSum.toFixed(0)).toLocaleString('en')}</br>
              Open projects amount  : ${Number(ProjectsSum.toFixed(0)).toLocaleString('en')}
            </div>

          </div>
        </div>

      </div>`;

    // console.log(this.domElement);
    this.renderCompleted();

  }
  //@ts-ignore
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


  // Backlog Old
  // orders amount : ${OrdersNotDelivered[0]}</br>
  // <table>
  //   <tr>
  //     <td class="${ styles.specTd }">Prototyping</td>
  //     <td class="${ styles.specTd }">Mprep</td>
  //     <td class="${ styles.specTd }">Clinical Builds</td>
  //     <td class="${ styles.specTd }">Manufacturing</td>
  //   </tr>
  //   <tr>
  //     <td> ${OrdersNotDelivered[1][0]}</td>
  //     <td> ${OrdersNotDelivered[1][1]}</td>
  //     <td> ${OrdersNotDelivered[1][2]}</td>
  //     <td> ${OrdersNotDelivered[1][3]}</td>
  //   </tr>
  // </table>
  //**************************** returns the full lists *************************/
  public getListItems(listname: string, listContainerName: string = null): void {

    // console.log('asking list items for', listname);
    this.ajaxCounter++;

    this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/web/lists/GetByTitle('${listname}')/Items?$top=5000`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        response.json().then((data) => {

          // console.log('list items for', listname, data);
          this.ajaxCounter--;
          if (listContainerName) {

            if (this.listsContainer[listContainerName]) {
              this.listsContainer[listContainerName] = this.listsContainer[listContainerName].concat(data.value)
            } else {
              this.listsContainer[listContainerName] = data.value;
            }

          } else {
            this.listsContainer[listname] = data.value;
          }

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


// <table>
// <tr>
//   <td>week1</td><td>week2</td><td>week3</td><td>week4</td><td>week5</td>
// </tr>
// <tr>
//   <td>${nextMonthInvoices[2][0]}</td>
//   <td>${nextMonthInvoices[2][1]}</td>
//   <td>${nextMonthInvoices[2][2]}</td>
//   <td>${nextMonthInvoices[2][3]}</td>
//   <td>${nextMonthInvoices[2][4]}</td>
// </tr>
// </table>

// 2023
// <table>
// <tr>
//   <td>type II</td>
//   <td>type III</td>
//   <td>type IV</td>
// </tr>
// <tr>
//   <td> ${monthlyQuotes[2][0]}%</td>
//   <td> ${monthlyQuotes[2][1]}%</td>
//   <td> ${monthlyQuotes[2][2]}%</td>
// </tr>
// </table>