/* eslint-disable @typescript-eslint/no-floating-promises */
import * as React from "react";
import styles from "./Demo.module.scss";
import { IFile, IResponseItem } from "./interface";
import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";

// import { PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';
import { Logger, LogLevel } from "@pnp/logging";

export interface IDemoProps {
  description: string;
  sp: SPFI;
}

export interface IDemoState {
  items: IFile[];
  errors: string[];
  userEmail: string;
}

export default class Demo extends React.Component<IDemoProps, IDemoState> {
  private currentUserEmail: string = '';
  LOG_SOURCE: string = 'Demo';
  LIBRARY_NAME: string = 'Documents';

  constructor(props: IDemoProps) {
    super(props);
    this.state = {
      items: [],
      errors: [],
      userEmail: ''
    };
  }

  public async componentDidMount(): Promise<void> {
    await this._getCurrentUserEmail();
    await this._readAllFilesSize(this.LIBRARY_NAME);
  }

  public render(): React.ReactElement<IDemoProps> {
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
            <p className="ms-font-l ms-fontColor-white">Current User Email: {this.state.userEmail}</p>
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
                                    <div className={styles.left}>{item.Title}</div>

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
            <button onClick={() => this._readAllFilesSize(this.LIBRARY_NAME)}>Refresh Files</button>
            <button onClick={() => this._batchUpdateItemTitles()}>Batch Update Item Titles</button>
          </div>
        </div>
      </div>
    );
  }

  private _getCurrentUserEmail = async (): Promise<void> => {
    try {
      const user = await this.props.sp.web.currentUser();
      this.currentUserEmail = user.Email;
      this.setState({ userEmail: user.Email });
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
      this.setState({ items });
    } catch (error) {
      Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(error)} - `, LogLevel.Error);
      this.setState({ errors: [...this.state.errors, error.message] });
    }
  };

  private _batchUpdateItemTitles = async (): Promise<void> => {
    try {
      const [batchedSP, execute] = this.props.sp.batched();
      const list = batchedSP.web.lists.getByTitle(this.LIBRARY_NAME);

      for (let i = 0; i < this.state.items.length; i++) {
        list.items.getById(this.state.items[i].Id).update({ Title: `${this.state.items[i].Name}-Updadfuvihhhhhhted` }).catch(async (error) => {
          if (error.message.includes('is locked for shared use')) {
            console.error(`File is locked for shared use: ${this.state.items[i].Name}`);
            const lockedByUser = await this._getLockedByUser(this.state.items[i].Name);
            const errorMessage = `File is locked: ${this.state.items[i].Name} by ${lockedByUser}`;
            this.setState(prevState => ({
              errors: [...prevState.errors, errorMessage]
            }));
          } else {
            throw error;
          }
        });
      }

      await execute();
      console.log('Batch update executed');
       // Refresh  items
      await this._readAllFilesSize(this.LIBRARY_NAME);
    } catch (error) {
      Logger.write(`Error batch updating item titles: ${JSON.stringify(error)}`, LogLevel.Error);
      this.setState({ errors: [...this.state.errors, error.message] });
    }
  };

  private _getLockedByUser = async (fileName: string): Promise<string> => {
    try {
      const file = this.props.sp.web.getFolderByServerRelativePath(this.LIBRARY_NAME).files.getByUrl(fileName);
      const user = await file.getLockedByUser();
      console.log(user);

      return user?.Email || 'Unknown user';
    } catch (error) {
      Logger.write(`Error getting locked by user for file ${fileName}: ${JSON.stringify(error)}`, LogLevel.Error);
      return 'Unknown user';
    }
  };

  private _getErrors(): JSX.Element | null {
    return this.state.errors.length > 0 ? (
      <div style={{ color: "orangered" }}>
        <div>Errors:</div>
        {this.state.errors.map((item, idx) => (
          <div key={idx}>{item}</div>
        ))}
      </div>
    ) : null;
  }
}