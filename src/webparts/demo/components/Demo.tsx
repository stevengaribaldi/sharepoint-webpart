import * as React from "react";
import styles from "./Demo.module.scss";

import { IFile, IResponseItem } from "./interface";

import { SPFI } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

export interface IAsyncAwaitPnPJsProps {
  description: string;
  sp: SPFI;
}

export interface IAsyncAwaitPnPJsState {
  items: IFile[];
  errors: string[];
}

export default class Demo extends React.Component<IAsyncAwaitPnPJsProps, IAsyncAwaitPnPJsState> {

  constructor(props: IAsyncAwaitPnPJsProps) {
    super(props);
    // set initial state
    this.state = {
      items: [],
      errors: []
    };
  }

  public componentDidMount(): void {
    // read all file sizes from Documents library
    this._readAllFilesSize("Documents");
  }

  public render(): React.ReactElement<IAsyncAwaitPnPJsProps> {
    // calculate total of file sizes
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
                <div className={`${styles.clear} ${styles.header}`}></div>
              </div>
              {this.state.items.map((item, idx) => (
                <div key={idx} className={styles.row}>
                  <div className={styles.left}>{item.Name}</div>
                  <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
                  <div className={styles.clear}></div>
                </div>
              ))}
              <div className={styles.row}>
                <div className={`${styles.clear} ${styles.header}`}></div>
                <div className={styles.left}>Total: </div>
                <div className={styles.right}>{(totalDocs / 1024).toFixed(2)}</div>
                <div className={`${styles.clear} ${styles.header}`}></div>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  private _readAllFilesSize = async (libraryName: string): Promise<void> => {
    try {
      const response: IResponseItem[] = await this.props.sp.web.lists
        .getByTitle(libraryName)
        .items.select("Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();

      // use map to convert IResponseItem[] into our internal object IFile[]
      const items: IFile[] = response.map((item: IResponseItem) => ({
        Id: item.Id,
        Title: item.Title,
        Size: item.File.Length,
        Name: item.FileLeafRef
      }));

      // set our Component´s State
      this.setState({ items });

      // intentionally set wrong query to see console errors...
      const failResponse: IResponseItem[] = await this.props.sp.web.lists
        .getByTitle(libraryName)
        .items.select("Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();
    } catch (error) {
      // set a new state conserving the previous state + the new error
      this.setState({ errors: [...this.state.errors, error] });
    }
  };

  private _getErrors() {
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
