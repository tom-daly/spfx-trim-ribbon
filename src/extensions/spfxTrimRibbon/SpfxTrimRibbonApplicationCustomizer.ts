import { override } from '@microsoft/decorators';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import "@pnp/polyfill-ie11";
import pnp, { PermissionKind } from "sp-pnp-js/lib/pnp";
import * as strings from 'SpfxTrimRibbonApplicationCustomizerStrings';

export interface ISpfxTrimRibbonApplicationCustomizerProperties {}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxTrimRibbonApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxTrimRibbonApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    return super.onInit().then(() => {
      pnp.setup({
        spfxContext: this.context,
        defaultCachingStore: "session",
        globalCacheDisable: false
      });
      this.trimRibbon();
    });
  }

  private trimRibbon(): void {
    pnp.sp.web
      .usingCaching()
      .currentUserHasPermissions(PermissionKind.EditListItems)
      .then(perms => {
        let suiteBar: HTMLElement = this.getSuiteBar();
        if (!suiteBar) return;
        if (!perms) {
          suiteBar.setAttribute("style", "display: none !important");
        }
      });
  }

  private getSuiteBar(): HTMLElement {
    return (
      document.getElementById("SuiteNavPlaceHolder") ||
      (document.getElementsByClassName("od-SuiteNav")[0] as HTMLElement)
    );
  }
}
