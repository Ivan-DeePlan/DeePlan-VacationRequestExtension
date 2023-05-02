import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  IListViewCommandSetListViewUpdatedParameters,
} from "@microsoft/sp-listview-extensibility";
import { override } from "@microsoft/decorators";
import { Constants } from "./Models/Constants";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IVacationRequestCommandSetProperties {
  // This is an example; replace with your own properties
  Create_Field: string;
  Update_Field: string;
}

const LOG_SOURCE: string = "VacationRequestCommandSet";

export default class VacationRequestCommandSet extends BaseListViewCommandSet<IVacationRequestCommandSetProperties> {
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized MatnasimFormsManagmentExtCommandSet");
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    //* Create Field Command
    const compareOneCommand: Command = this.tryGetCommand("Create_Field");
    //* Update Field Command visible only when selected item
    const compareSecondCommand: Command = this.tryGetCommand("Update_Field");

    //* Get current library name from pageContext
    let LibraryName = this.context.pageContext.list.title;
    if (Constants.LIBRARY_NAME.indexOf(LibraryName) !== -1) {
      compareOneCommand.visible = true;
      if (compareSecondCommand) {
        compareSecondCommand.visible = event.selectedRows?.length === 1;
      }
      require("./ExtStyle/ExtStyle.css");
    } else {
      compareOneCommand.visible = false;
      compareSecondCommand.visible = false;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const webUri = this.context.pageContext.web.absoluteUrl;

    switch (event.itemId) {
      //*Go to create page
      case "Create_Field":
        window.location.href = webUri + Constants.newListView;
        break;

      //*Go to update page with selected item Id
      case "Update_Field":
        var ItemID = event.selectedRows[0].getValueByName("ID").toString();
        window.location.href =
          webUri + Constants.EditListView + "?FormID=" + ItemID;
        break;
      default:
        throw new Error("Unknown command");
    }
  }
}
