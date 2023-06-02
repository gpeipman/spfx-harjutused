import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
//import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import styles from './BirthdaysWebPart.module.scss';
import * as strings from 'BirthdaysWebPartStrings';

import { Birthday } from './Birthday';
import { IBirthdays } from './IBirthdays';
import { IBirthdaysWebPartProps } from './IBirthdaysWebPartProps'

export default class BirthdaysWebPart extends BaseClientSideWebPart<IBirthdaysWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.title}">${this.properties.title}</div>
      <div id="birthdaysContainer" />
    `;

    if(!this.properties.listName) {
      return;
    }

    this._renderBirthdaysAsync();
  }

  // Küsi sünnipäevad SharePointist, teisenda need Birthday tüüpi objektideks,
  // sorteeri sünnipäevad ära ja kirjuta need ekraanile
  private _renderBirthdaysAsync(): void {
    // Küsi sünnipäevad
    this._getBirthdays()
        .then((response) => {
          // SharePointist tulevad Birthday laadsed objektid, mis pole päris Birthday siiski.
          // Teisendame need päris Birthday tüüpi objektideks
          let birthdays = Birthday.fromRandomObjects(response.value)

          // SharePoint ei tea kuidas sünnipäevi sorteerida. Sorteerime need ise ära.
          birthdays = Birthday.sort(birthdays);

          // Laseme sünnipäevad välja trükkida
          this._renderBirthdays(birthdays);
        })
        .catch((ex) => { console.error(ex); });
  }

  private _renderBirthdays(items: Birthday[]): void {    
    let html: string = '';
    let lastDay: number = 0;

    html += '<table>';

    items.forEach((item: Birthday) => {
      if(!item.isAtCurrentMonth()) {
        return;
      }

      // Sama päeva sünnipäevade kuupäeva kirjutame välja ainult ühe korra
      // lastDay muutuja näitab, millise kuupäeva kirjutasime välja viimati
      let date = '';
      if(item.Day !== lastDay) {
        date = item.formatAsDayAndMonth();
        lastDay = item.Day;
      }

      // Kui sünnipäev on täna, siis kirjutame selle välja paksus kirjas
      let todayStyle = '';
      if(item.isToday()) {
        todayStyle = styles.today;
      }

      html += `
          <tr class="${todayStyle}">
            <td>${date}</td>
            <td>${escape(item.Title)}</td>
          </tr>
        `;

        console.log(typeof(item.Month));
    });
  
    html += '</table>';

    const listContainer: Element = this.domElement.querySelector('#birthdaysContainer');
    listContainer.innerHTML = html;
  }

  private _getBirthdays() : Promise<IBirthdays> {
    let url = this.context.pageContext.web.absoluteUrl;
    url += "/_api/web/lists/getByTitle('Birthdays')/items?";
    url += "$select=Title,Month,Day";

    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1)
            .then((response: SPHttpClientResponse) => {
              return response.json();
            })
            .catch((ex) => { console.error(ex); });
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel,
                  onGetErrorMessage: this.validateListName.bind(this),
                  deferredValidationTime: 500
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private async validateListName(value: string): Promise<string> {
    if (value === null || value.length === 0) {
      return "";
    }

    try {
      const response = await this.context.spHttpClient.get(
        this.context.pageContext.web.absoluteUrl +
          `/_api/web/lists/getByTitle('${escape(value)}')?$select=Id`,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        return "";
      } else if (response.status === 404) {
        return `List '${escape(value)}' doesn't exist in the current site`;
      } else {
        return `Error: ${response.statusText}. Please try again`;
      }
    } catch (error) {
      return error.message;
    }
  }
}
