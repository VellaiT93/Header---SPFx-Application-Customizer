import { override } from '@microsoft/decorators';
import {
  Log, Environment,
  EnvironmentType
} from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

import * as strings from 'HeaderApplicationCustomizerStrings';
import styles from './HeaderStyles.module.scss';

const LOG_SOURCE: string = 'HeaderApplicationCustomizer';

export interface ISPLists {
  value: ISPList[];
}
export interface ISPList {
  Title: string;
  Show: string;
}

export interface IHeaderApplicationCustomizerProperties {
  testMessage: string;
}

const Pictures = {
  emergencyBlack: require('./pics/emergency_icon_black.png'),
  emergencyRed: require('./pics/emergency_icon_red.png'),
  contactClosed: require('./pics/contact_icon_closed.png'),
  contactOpen: require('./pics/contact_icon_open.png'),
  tippsBlack: require('./pics/tipps_black.png'),
  tippsYellow: require('./pics/tipps_yellow.png'),
}

/** Application Customizer: Top Placeholder */
export default class HeaderApplicationCustomizer extends BaseApplicationCustomizer<IHeaderApplicationCustomizerProperties> {

  private topPlaceholder: PlaceholderContent | undefined;
  private queryString: String = "http://sp2019server/sites/it/_api/web/lists/GetByTitle('HeaderText')/Items";

  // Render function of text
  renderText(items: ISPList[]): void {
    const container: HTMLElement = this.topPlaceholder.domElement.querySelector('[name="placeHolder"]') as HTMLElement;
    const content: HTMLElement = document.createElement('div');
    content.setAttribute('id', `${styles.content}`);

    container.appendChild(content);

    let textHolder: String = ``;

    // Set the Title as text
    items.forEach((item: ISPList) => {
      if (item.Show.toString() === 'true') {
        textHolder += `<div class='${styles.text}'>${item.Title}</div>`;
      }
    });

    while (content.offsetWidth < container.offsetWidth) {
      content.innerHTML += textHolder;
    }

    content.style.width = content.offsetWidth + 'px';

    content.innerHTML += content.innerHTML;
  }

  // Get list items from SP list
  async getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(`${this.queryString}`,
      SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

    if (!this.topPlaceholder) {
      console.error('There is an error!');
      return;
    } else {
      if (this.topPlaceholder.domElement) {
        this.topPlaceholder.domElement.innerHTML = `
          <div name='topPlaceholder' id='${styles.topPlaceholder}'>
            <div name='menu' id='${styles.menu}'>

              <div class='${styles.dropdown}'>
                <button class='${styles.dropdownButton}'>Dropdown</button>
                <div class='${styles.dropdownContent}'>
                  <a href="#">Link 1</a>
                  <a href="#">Link 2</a>
                  <a href="#">Link 3</a>
                </div>
              </div>

              <div class='${styles.dropdown}'>
                <button class='${styles.dropdownButton}'>Dropdown</button>
                <div class='${styles.dropdownContent}'>
                  <a href="#">Link 1</a>
                  <a href="#">Link 2</a>
                  <a href="#">Link 3</a>
                </div>
              </div>

              <div class='${styles.dropdown}'>
                <button class='${styles.dropdownButton}'>Dropdown</button>
                <div class='${styles.dropdownContent}'>
                  <a href="#">Link 1</a>
                  <a href="#">Link 2</a>
                  <a href="#">Link 3</a>
                </div>
              </div>

              <div class='${styles.dropdown}'>
                <button class='${styles.dropdownButton}'>Dropdown</button>
                <div class='${styles.dropdownContent}'>
                  <a href="#">Link 1</a>
                  <a href="#">Link 2</a>
                  <a href="#">Link 3</a>
                </div>
              </div>

            </div>
            <div name='placeHolder' id='${styles.placeHolder}'></div>
            <div name='right' id='${styles.right}'>

              <div class='${styles.emergency}'>
                <img id='${styles.emergency_black}' src='${Pictures.emergencyBlack}' />
                <img id='${styles.emergency_red}'
                  src='${Pictures.emergencyRed}'
                />
              </div>

              <div class='${styles.emergency}'>
                <img id='${styles.emergency_black}' src='${Pictures.contactClosed}' />
                <img
                  id='${styles.emergency_red}'
                  src='${Pictures.contactOpen}'
                />
              </div>

              <div class='${styles.emergency}'>
                <img id='${styles.emergency_black}' src='${Pictures.tippsBlack}' />
                <img
                  id='${styles.emergency_red}'
                  src='${Pictures.tippsYellow}'
                />
              </div>

            </div>
          </div>
        `;

        if (Environment.type === EnvironmentType.SharePoint) {
          this.getListData().then((response) => {
            console.log(response);
            this.renderText(response.value)
          });
        }
      }
    }

    return Promise.resolve();
  }
}
