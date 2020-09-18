import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'PolymorphicCraWebPartStrings';

import styles from './PolymorphicCraWebPart.module.scss';

export interface IPolymorphicCraWebPartProps {
  webAppUrl: string;
}

const IFRAME_ID = "Polymorphic-CRA-IFRAME";

export default class PolymorphicCraWebPart extends BaseClientSideWebPart<IPolymorphicCraWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <iframe id="${IFRAME_ID}" class="${styles.PolymorphicCRAIframe}" src="${this.properties.webAppUrl}"></iframe>
    `;

    try {
      const target: HTMLIFrameElement = document.querySelector(`#${IFRAME_ID}`);
      target.addEventListener('load', () => {
        const payload = {
          web: this.context.pageContext.web,
        };
        target.contentWindow.postMessage(payload, this.properties.webAppUrl);
      });
    }
    catch (error) {
      console.log(`Unable to send message to iframe`, error);
    }
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
                PropertyPaneTextField('webAppUrl', {
                  label: strings.WebAppUrlFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
