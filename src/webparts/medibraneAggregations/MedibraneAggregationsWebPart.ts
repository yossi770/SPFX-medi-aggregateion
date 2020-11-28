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

  public render(): void {
/*
    this.getListItems('Quotes').then((items)=>{});
    this.getListItems('Orders').then((items)=>{});
    this.getListItems('Projects').then((items)=>{});
*/

    this.getListItemsShhh('Quotes');
    this.getListItemsShhh('Orders');
    this.getListItemsShhh('Projects');

    
  }

  ajaxCounter:number = 0;
  listsContainer:{} = {};
  
  public buildHtml(){
    let quotesTotal = 0;
    let oredersTotal = 0;
    let projTotal = 0;

    let forche = (arr:[], fName:string) => {
      let t = 0
      for (let i = 0; i < arr.length; i++) {
        const item = arr[i];
        t += item[fName]
      }
      return t;
    }

    quotesTotal = forche(this.listsContainer['Quotes'], 'Quota_x0020_amount')
    oredersTotal = forche(this.listsContainer['Orders'], 'Order_x0020_Amount')
    projTotal = forche(this.listsContainer['Projects'], 'Order_x0020_Amount')



    this.domElement.innerHTML = `
      <div class="${ styles.medibraneAggregations }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>
              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>

              <h2>quotes :  ${quotesTotal}</h2>
              <h2>orders :  ${oredersTotal}</h2>
              <h2>projs :  ${projTotal}</h2>

            </div>
          </div>
        </div>
      </div>`;

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


  public getListItems(listname:string): Promise<{}[]> {
    console.log('asking list items for', listname);
    this.ajaxCounter++;

    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl + 
      `/_api/web/lists/GetByTitle('${listname}')/Items`, SPHttpClient.configurations.v1)
          .then((response: SPHttpClientResponse) => {
              let items = response.json()['value'];
              console.log('list items for', listname, items);
              this.ajaxCounter--;
              if (this.ajaxCounter == 0) {
                
              }
              return items;
          });
    }

    public getListItemsShhh(listname:string): void {
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
