import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import {AppInsights} from "applicationinsights-js";
import * as strings from 'SpFxAcAppInsightsApplicationCustomizerStrings';

const LOG_SOURCE: string = 'SpFxAcAppInsightsApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpFxAcAppInsightsApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpFxAcAppInsightsApplicationCustomizer
  extends BaseApplicationCustomizer<ISpFxAcAppInsightsApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }


    /* update with YOUR App Insights key: */
    let appInsightsKey: string = "bd755e10-2e26-4580-8d58-446c0dd329fe";

    AppInsights.downloadAndSetup({ instrumentationKey: appInsightsKey });

    // simple usage - all params will be derived..
    AppInsights.trackPageView();

    console.log(`OnInit: Called trackPageView().`);

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    return Promise.resolve();
  }
}
