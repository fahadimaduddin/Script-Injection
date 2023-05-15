import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'ScriptInjectionApplicationCustomizerStrings';

const LOG_SOURCE: string = 'ScriptInjectionApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IScriptInjectionApplicationCustomizerProperties {
  // This is an example; replace with your own property
  cssurl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class ScriptInjectionApplicationCustomizer
  extends BaseApplicationCustomizer<IScriptInjectionApplicationCustomizerProperties> {

  public onInit(): Promise<void> {
    let cssFileUrl:string="https://fahadimaduddin.sharepoint.com/SiteAssets/Asset/style.css";
    let scriptFileUrl:string="https://fahadimaduddin.sharepoint.com/SiteAssets/Asset/index.js";
    let jQueryFileUrl:string="https://fahadimaduddin.sharepoint.com/SiteAssets/Asset/jquery-1.7.2.js";
    let head:any = document.getElementsByTagName("head")[0] || document.documentElement;
    console.log("Demo " + head);
    if(cssFileUrl){
      let linkTag: HTMLLinkElement = document.createElement("link");
      linkTag.href = cssFileUrl;
      linkTag.type = "text/css";
      linkTag.rel = "stylesheet";
      head.insertAdjacentElement("beforeEnd",linkTag);
    }
    if(scriptFileUrl){
      let jQueryRef: HTMLScriptElement = document.createElement("script");
      jQueryRef.src = jQueryFileUrl;
      jQueryRef.type = "text/javascript";
      head.insertAdjacentElement(jQueryRef);

      let CustScript: HTMLScriptElement = document.createElement("script");
      CustScript.src = scriptFileUrl;
      CustScript.type = "text/javascript";
      head.insertAdjacentElement(CustScript);
    }
    return Promise.resolve();
  }
}
