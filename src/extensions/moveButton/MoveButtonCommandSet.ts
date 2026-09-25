import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs,
  type RowAccessor
} from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { MoveFolderPicker } from './components/MoveFolderPicker';
import { SitePagesService } from './services/SitePagesService';

const LOG_SOURCE: string = 'MoveButtonCommandSet';

export interface IMoveButtonCommandSetProperties {}

export default class MoveButtonCommandSet extends BaseListViewCommandSet<IMoveButtonCommandSetProperties> {

  private _container: HTMLElement | null = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized MoveButtonCommandSet');

    const moveCommand: Command = this.tryGetCommand('MOVE_PAGE');
    if (moveCommand) moveCommand.visible = false;

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'MOVE_PAGE':
        this._openPicker();
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private _openPicker(): void {
    const selectedRows = this.context.listView.selectedRows;
    const list = this.context.listView.list;
    if (!selectedRows || selectedRows.length === 0 || !list) return;

    const selectedFiles = selectedRows
      .filter(row => !MoveButtonCommandSet._isFolder(row))
      .map(row => ({
        name: row.getValueByName('FileLeafRef') as string,
        serverRelativeUrl: row.getValueByName('FileRef') as string
      }));
    if (selectedFiles.length === 0) return;

    const service = new SitePagesService(
      this.context.spHttpClient,
      this.context.pageContext.web.absoluteUrl,
      list.serverRelativeUrl
    );

    this._closePicker();
    this._container = document.createElement('div');
    document.body.appendChild(this._container);

    const element = React.createElement(MoveFolderPicker, {
      service,
      libraryTitle: list.title,
      selectedFiles,
      onDismiss: () => this._closePicker(),
      onMoveComplete: () => {
        this._closePicker();
        window.location.reload();
      }
    });

    ReactDOM.render(element, this._container);
  }

  private _closePicker(): void {
    if (this._container) {
      ReactDOM.unmountComponentAtNode(this._container);
      this._container.remove();
      this._container = null;
    }
  }

  public onDispose(): void {
    this._closePicker();
    super.onDispose();
  }

  private static _isFolder(row: RowAccessor): boolean {
    return String(row.getValueByName('FSObjType')) === '1';
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    const moveCommand: Command = this.tryGetCommand('MOVE_PAGE');
    if (moveCommand) {
      const selectedRows = this.context.listView.selectedRows ?? [];
      moveCommand.visible = selectedRows.length >= 1 && !selectedRows.some(MoveButtonCommandSet._isFolder);
    }
    this.raiseOnChange();
  }
}
