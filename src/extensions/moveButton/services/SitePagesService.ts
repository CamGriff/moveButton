import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

export interface IFolder {
  Name: string;
  ServerRelativeUrl: string;
}

// System folders in Site Pages that pages should never be moved into
const EXCLUDED_FOLDERS: string[] = ['Forms', 'Templates'];

export class SitePagesService {
  constructor(
    private readonly spHttpClient: SPHttpClient,
    private readonly webUrl: string,
    public readonly libraryRootUrl: string
  ) {}

  public async getFolders(serverRelativeUrl: string): Promise<IFolder[]> {
    const filter = EXCLUDED_FOLDERS.map(name => `Name ne '${name}'`).join(' and ');
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      `${this.webUrl}/_api/web/GetFolderByServerRelativePath(decodedurl='${SitePagesService._odataString(serverRelativeUrl)}')/Folders` +
        `?$select=Name,ServerRelativeUrl&$filter=${encodeURIComponent(filter)}&$orderby=Name`,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) throw new Error(await SitePagesService._errorMessage(response));

    const data = await response.json();
    return data.value as IFolder[];
  }

  /** Moves a file without overwriting; fails if a file with the same name exists in the target. */
  public async movePage(sourceUrl: string, targetFolderUrl: string, fileName: string): Promise<void> {
    const origin = new URL(this.webUrl).origin;

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      `${this.webUrl}/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=false`,
      SPHttpClient.configurations.v1,
      {
        body: JSON.stringify({
          srcPath: { DecodedUrl: `${origin}${sourceUrl}` },
          destPath: { DecodedUrl: `${origin}${targetFolderUrl}/${fileName}` },
          options: { KeepBoth: false, ResetAuthorAndCreatedOnCopy: false, ShouldBypassSharedLocks: false }
        })
      }
    );

    if (!response.ok) throw new Error(await SitePagesService._errorMessage(response));
  }

  // Escapes a value for use inside a single-quoted OData string literal in a URL
  private static _odataString(value: string): string {
    return encodeURIComponent(value.replace(/'/g, "''"));
  }

  private static async _errorMessage(response: SPHttpClientResponse): Promise<string> {
    try {
      const data = await response.json();
      const error = data['odata.error'] ?? data.error;
      const message = error?.message?.value ?? error?.message;
      if (message) return message;
    } catch {
      // Response body was not JSON
    }
    return `Request failed with status ${response.status}`;
  }
}
