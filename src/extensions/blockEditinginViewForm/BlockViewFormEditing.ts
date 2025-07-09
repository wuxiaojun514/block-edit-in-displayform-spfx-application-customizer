import { Log } from "@microsoft/sp-core-library";
import * as jQuery from "jquery";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

import * as strings from "BlockViewFormEditingApplicationCustomizerStrings";

const LOG_SOURCE: string = "BlockViewFormEditingApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IBlockViewFormEditingProperties {
  // put the regex of the list/document library you want to block view form, it will block all lists/libraries if this parameter is empty
  // e.g.    \/Lists\/(list1|my%20list)\/  
  blockAddressRegex: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class BlockViewFormEditingApplicationCustomizer extends BaseApplicationCustomizer<IBlockViewFormEditingProperties> {
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    const blockAddressRegex = this.properties.blockAddressRegex;
    // eslint-disable-next-line no-useless-escape
    let regex = RegExp("\/(Lists|Forms)\/"); //It will block view form on all lists/libraries by default if you don't specify the block address
    if (blockAddressRegex && blockAddressRegex.length>1 && this.isValidRegex(blockAddressRegex)) {
      regex = RegExp(blockAddressRegex);
    }

    if (regex.test(window.location.pathname)) {
      setInterval(this.blockViewFormEditting, 2000);
      Log.info(LOG_SOURCE,
        "block click event on display form field control to force end user use edit form"
      );
    }

    return Promise.resolve();
  }

  private blockViewFormEditting(): void {
    //new dialog display form
    if (jQuery(".sp-itemDialog .sp-itemDialogForm").length > 0) {
      if (jQuery('.sp-itemDialogHeader button[title="Edit all"]').length > 0) {
        $(".ReactFieldEditor-core--display").on("click", (event) => {
          event.preventDefault(); // Prevent the default action
          event.stopPropagation(); // Stop the event from propagating
        });
      }
    }

    //old panel display form
    if (jQuery(".list-form-container-root .list-form-client").length > 0) {
      if (
        jQuery('.list-form-container-root button[name="Edit all"]').length > 0
      ) {
        $(".ReactFieldEditor-core--display").on("click", (event) => {
          event.preventDefault(); // Prevent the default action
          event.stopPropagation(); // Stop the event from propagating
        });
      }
    }
  }

  private isValidRegex(pattern: string): boolean {
    try {
      RegExp(pattern);
      return true;
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    } catch (ex) {
      return false;
    }
  }
}
