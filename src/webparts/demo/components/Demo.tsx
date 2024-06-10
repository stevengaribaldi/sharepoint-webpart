import * as React from "react";
import styles from "./Demo.module.scss";
import { IFile, IResponseItem } from "./interface";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
// import { PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';
import { Logger, LogLevel } from "@pnp/logging";

export interface IAsyncAwaitPnPJsProps {
  description: string;
  sp: SPFI;
}

export interface IAsyncAwaitPnPJsState {
  items: IFile[];
  errors: string[];
}

export default class Demo extends React.Component<IAsyncAwaitPnPJsProps, IAsyncAwaitPnPJsState> {
  private currentUserEmail: string = '';
  LOG_SOURCE: string = 'Demo';

  constructor(props: IAsyncAwaitPnPJsProps) {
    super(props);
    this.state = {
      items: [],
      errors: []
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._getCurrentUserEmail();
    await this._readAllFilesSize("Documents");
  }

  public render(): React.ReactElement<IAsyncAwaitPnPJsProps> {
    const totalDocs: number = this.state.items.length > 0
      ? this.state.items.reduce<number>((acc: number, item: IFile) => {
          return acc + Number(item.Size);
        }, 0)
      : 0;

    return (
      <div className={styles.demo}>
        <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
          <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
            <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
            <div>{this._getErrors()}</div>
            <p className="ms-font-l ms-fontColor-white">List of documents:</p>
            <div>
              <div className={styles.row}>
                <div className={styles.left}>Name</div>
                <div className={styles.right}>Size (KB)</div>
                <div className={`${styles.clear} ${styles.header}`} />
              </div>
              {this.state.items.map((item, idx) => (
                <div key={idx} className={styles.row}>
                  <div className={styles.left}>{item.Name}</div>
                  <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
                  <div className={styles.clear} />
                </div>
              ))}
              <div className={styles.row}>
                <div className={`${styles.clear} ${styles.header}`} />
                <div className={styles.left}>Total: </div>
                <div className={styles.right}>{(totalDocs / 1024).toFixed(2)}</div>
                <div className={`${styles.clear} ${styles.header}`} />
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _getCurrentUserEmail = async (): Promise<void> => {
    try {
      const user = await this.props.sp.web.currentUser();
      this.currentUserEmail = user.Email;
      Logger.write(`Current user email: ${this.currentUserEmail}`, LogLevel.Info);
    } catch (error) {
      Logger.write(`Error getting current user email: ${JSON.stringify(error)}`, LogLevel.Error);
      this.setState({ errors: [...this.state.errors, error.message] });
    }
  };

  private _readAllFilesSize = async (libraryName: string): Promise<void> => {
    try {
      const response: IResponseItem[] = await this.props.sp.web.lists
        .getByTitle(libraryName)
        .items.select("Id", "Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();

      const items: IFile[] = response.map((item: IResponseItem) => ({
        Id: item.Id,
        Title: item.Title,
        Size: item.File.Length,
        Name: item.FileLeafRef
      }));

      console.log(items);
    } catch (error) {
      Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(error)} - `, LogLevel.Error);
    }
  };

  private _getErrors(): JSX.Element | null {
    return this.state.errors.length > 0 ? (
      <div style={{ color: "orangered" }}>
        <div>Errors:</div>
        {this.state.errors.map((item, idx) => (
          <div key={idx}>{JSON.stringify(item)}</div>
        ))}
      </div>
    ) : null;
  }
}
