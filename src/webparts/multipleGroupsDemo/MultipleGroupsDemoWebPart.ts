import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './MultipleGroupsDemoWebPart.module.scss';
import * as strings from 'MultipleGroupsDemoWebPartStrings';

export interface IMultipleGroupsDemoWebPartProps {
  description: string;
  productName: string;
  isCertified: boolean;
}

export default class MultipleGroupsDemoWebPart extends BaseClientSideWebPart<IMultipleGroupsDemoWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.multipleGroupsDemo }">
        <div class="${ styles.container }">
          <div class="${ styles.row }">
            <div class="${ styles.column }">
              <span class="${ styles.title }">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle }">Customize SharePoint experiences using Web Parts.</p>
              <p class="${ styles.description }">${escape(this.properties.description)}</p>

              <p class="${ styles.description }">${escape(this.properties.productName)}</p>
              <p class="${ styles.description }">${(this.properties.isCertified)}</p>

              <a href="https://aka.ms/spfx" class="${ styles.button }">
                <span class="${ styles.label }">Learn more</span>
              </a>
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
            description: 'Page #1'
          },
          groups: [
            {
              groupName: 'First Group',
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: 'Product Name'
                }) 
              ]
            },
            {
              groupName: 'Second Group',
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: 'Is certified?'
                })
              ]
            }
          ],
          displayGroupsAsAccordion: true
        },
        {
          header: {
            description: 'Page #2'
          },
          groups: [
            {
              groupName: 'First Group',
              groupFields: [
                PropertyPaneTextField('productName', {
                  label: 'Product Name'
                }) 
              ]
            },
            {
              groupName: 'Second Group',
              groupFields: [
                PropertyPaneToggle('isCertified', {
                  label: 'Is certified?'
                })
              ]
            }
          ],
          displayGroupsAsAccordion: false
        }
      ]
    };
  }
}
