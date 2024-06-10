// import * as React from 'react';
// import styles from './Demo.module.scss';
// import { IDemoProps } from './IDemoProps';

// // import interfaces
// import { IFile, IResponseItem } from "./interface";

// import { Caching } from "@pnp/queryable";
// import { getSP } from "../pnpjsConfig";
// import { SPFI, spfi } from "@pnp/sp";
// import { Logger, LogLevel } from "@pnp/logging";
// //@ts-ignore
// import { IItemUpdateResult } from "@pnp/sp/items";
// import { Label, PrimaryButton } from '@microsoft/office-ui-fabric-react-bundle';

// export interface IDemoProps {
//   description: string;
// }

// export interface IIPnPjsExampleState {
//   items: IFile[];
//   errors: string[];
// }

// export default class Demo extends React.Component<IDemoProps, IIPnPjsExampleState> {
//   private LOG_SOURCE = "🅿PnPjsExample";
//   private LIBRARY_NAME = "Documents";
//   private _sp: SPFI;

//   constructor(props: IDemoProps) {
//     super(props);
//     // set initial state
//     this.state = {
//       items: [],
//       errors: []
//     };
//     this._sp = getSP();
//   }

//   public componentDidMount(): void {
//     // read all file sizes from Documents library
//     this._readAllFilesSize();
//   }

//   public render(): React.ReactElement<IDemoProps> {
//     try {
//       // calculate total of file sizes
//       const totalDocs: number = this.state.items.length > 0
//         ? this.state.items.reduce<number>((acc: number, item: IFile) => {
//           return (acc + Number(item.Size));
//         }, 0)
//         : 0;
//       return (
//         <div className={styles.demo}>
//           <Label>Welcome to PnP JS Version 3 Demo!</Label>
//           <PrimaryButton onClick={this._updateTitles}>Update Item Titles</PrimaryButton>
//           <Label>List of documents:</Label>
//           <table width="100%">
//             <tbody>
//               <tr>
//                 <td><strong>Title</strong></td>
//                 <td><strong>Name</strong></td>
//                 <td><strong>Size (KB)</strong></td>
//               </tr>
//               {this.state.items.map((item, idx) => {
//                 return (
//                   <tr key={idx}>
//                     <td>{item.Title}</td>
//                     <td>{item.Name}</td>
//                     <td>{(item.Size / 1024).toFixed(2)}</td>
//                   </tr>
//                 );
//               })}
//               <tr>
//                 <td></td>
//                 <td><strong>Total:</strong></td>
//                 <td><strong>{(totalDocs / 1024).toFixed(2)}</strong></td>
//               </tr>
//             </tbody>
//           </table>
//         </div >
//       );
//     } catch (err) {
//       Logger.write(`${this.LOG_SOURCE} (render) - ${JSON.stringify(err)} - `, LogLevel.Error);
//     }
//     return <div></div>;
//   }

//   private _readAllFilesSize = async (): Promise<void> => {
//     try {
//       // do PnP JS query, some notes:
//       //   - .expand() method will retrieve Item.File item but only Length property
//       //   - .get() always returns a promise
//       //   - await resolves promises making your code act synchronous, ergo Promise<IResponseItem[]> becomes IResponse[]

//       //Extending our sp object to include caching behavior, this modification will add caching to the sp object itself
//       //this._sp.using(Caching("session"));

//       //Creating a new sp object to include caching behavior. This way our original object is unchanged.
//       const spCache = spfi(this._sp).using(Caching({ store: "session" }));

//       const response: IResponseItem[] = await spCache.web.lists
//         .getByTitle(this.LIBRARY_NAME)
//         .items
//         .select("Id", "Title", "FileLeafRef", "File/Length")
//         .expand("File/Length")();

//       // use map to convert IResponseItem[] into our internal object IFile[]
//       const items: IFile[] = response.map((item: IResponseItem) => {
//         return {
//           Id: item.Id,
//           Title: item.Title || "Unknown",
//           Size: item.File?.Length || 0,
//           Name: item.FileLeafRef
//         };
//       });

//       // Add the items to the state
//       this.setState({ items });
//     } catch (err) {
//       Logger.write(`${this.LOG_SOURCE} (_readAllFilesSize) - ${JSON.stringify(err)} - `, LogLevel.Error);
//     }
//   }

//   private _updateTitles = async (): Promise<void> => {
//     try {
//       //Will create a batch call that will update the title of each item
//       //  in the library by adding `-Updated` to the end.
//       const [batchedSP, execute] = this._sp.batched();

//       //Clone items from the state
//       const items = JSON.parse(JSON.stringify(this.state.items));

//       const res: IItemUpdateResult[] = [];

//       for (let i = 0; i < items.length; i++) {
//         // you need to use .then syntax here as otherwise the application will stop and await the result
//         batchedSP.web.lists
//           .getByTitle(this.LIBRARY_NAME)
//           .items
//           .getById(items[i].Id)
//           .update({ Title: `${items[i].Name}-Updated` })
//           .then(r => res.push(r));
//       }
//       // Executes the batched calls
//       await execute();

//       // Results for all batched calls are available
//       for (let i = 0; i < res.length; i++) {
//         //If the result is successful update the item
//         //NOTE: This code is over simplified, you need to make sure the Id's match
//         const item = await res[i].item.select<{ Id: number, Title: string }>("Id, Title");
//         items[i].Name = item.Title;
//       }

//       //Update the state which rerenders the component
//       this.setState({ items });
//     } catch (err) {
//       Logger.write(`${this.LOG_SOURCE} (_updateTitles) - ${JSON.stringify(err)} - `, LogLevel.Error);
//     }
//   }
// }


// import * as React from 'react';
// import * as ReactDom from 'react-dom';
// import { Version } from '@microsoft/sp-core-library';
// import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
// import {
//   IPropertyPaneConfiguration,
//   // PropertyPaneTextField
// } from '@microsoft/sp-property-pane';


// import * as strings from 'DemoWebPartStrings';
// import Demo from './Demo';
// import { IDemoProps } from './IDemoProps';


// import { spfi, SPFI, SPFx } from "@pnp/sp";

// export interface IDemoWebPartProps {
//   description: string;
// }

// export default class DemoWebPart extends BaseClientSideWebPart<IDemoWebPartProps> {
//   private sp: SPFI;

//   // // https://github.com/SharePoint/PnP-JS-Core/wiki/Using-sp-pnp-js-in-SharePoint-Framework
//   public async onInit(): Promise<void> {
//     await super.onInit();

//     this.sp = spfi().using(SPFx(this.context));
//   }

//   public render(): void {
//     const element: React.ReactElement<IDemoProps> = React.createElement(
//       Demo,
//       {
//         description: this.properties.description,
//         sp: this.sp
//       }
//     );

//     ReactDom.render(element, this.domElement);
//   }

//   protected get dataVersion(): Version {
//     return Version.parse('1.0');
//   }

//   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
//     return {
//       pages: [
//         {
//           header: {
//             description: strings.PropertyPaneDescription
//           },
//           groups: [
//             {
//               groupName: strings.BasicGroupName,
//               groupFields: [
//                 PropertyPaneTextField('description', {
//                   label: strings.DescriptionFieldLabel
//                 })
//               ]
//             }
//           ]
//         }
//       ]
//     };
//   }
// }

//
//
// // import * as React from "react";
// import styles from "./Demo.module.scss";

// import { IFile, IResponseItem } from "./interface";

// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

// export interface IAsyncAwaitPnPJsProps {
//   description: string;
//   sp: SPFI;
// }

// export interface IAsyncAwaitPnPJsState {
//   items: IFile[];
//   errors: string[];
// }

// export default class Demo extends React.Component<IAsyncAwaitPnPJsProps, IAsyncAwaitPnPJsState> {

//   constructor(props: IAsyncAwaitPnPJsProps) {
//     super(props);
//     // set initial state
//     this.state = {
//       items: [],
//       errors: []
//     };
//   }

//  componentDidMount(): void {
//     this._readAllFilesSize("Documents").catch(error => console.error(error));
//   }

//   public render(): React.ReactElement<IAsyncAwaitPnPJsProps> {
//     // calculate total of file sizes
//     const totalDocs: number = this.state.items.length > 0
//       ? this.state.items.reduce<number>((acc: number, item: IFile) => {
//           return acc + Number(item.Size);
//         }, 0)
//       : 0;

//     return (
//       <div className={styles.demo}>
//         <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
//           <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
//             <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
//             <div>{this._getErrors()}</div>
//             <p className="ms-font-l ms-fontColor-white">List of documents:</p>
//             <div>
//               <div className={styles.row}>
//                 <div className={styles.left}>Name</div>
//                 <div className={styles.right}>Size (KB)</div>
//                 <div className={`${styles.clear} ${styles.header}`}></div>
//               </div>
//               {this.state.items.map((item, idx) => (
//                 <div key={idx} className={styles.row}>
//                   <div className={styles.left}>{item.Name}</div>
//                   <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
//                   <div className={styles.clear}></div>
//                 </div>
//               ))}
//               <div className={styles.row}>
//                 <div className={`${styles.clear} ${styles.header}`}></div>
//                 <div className={styles.left}>Total: </div>
//                 <div className={styles.right}>{(totalDocs / 1024).toFixed(2)}</div>
//                 <div className={`${styles.clear} ${styles.header}`}></div>
//               </div>
//             </div>
//           </div>
//         </div>

//       </div>
//     );
//   }

//   private _readAllFilesSize = async (libraryName: string): Promise<void> => {
//     try {
//       const response: IResponseItem[] = await this.props.sp.web.lists
//         .getByTitle(libraryName)
//         .items.select("Title", "FileLeafRef", "File/Length")
//         .expand("File/Length")();

//       // use map to convert IResponseItem[] into our internal object IFile[]
//       const items: IFile[] = response.map((item: IResponseItem) => ({
//         Id: item.Id,
//         Title: item.Title,
//         Size: item.File.Length,
//         Name: item.FileLeafRef
//       }));

//       // set our Component´s State
//       this.setState({ items });

//       // intentionally set wrong query to see console errors...
//       // const failResponse: IResponseItem[] = await this.props.sp.web.lists
//       //   .getByTitle(libraryName)
//       //   .items.select("Title", "FileLeafRef", "File/Length")
//       //   .expand("File/Length")();
//     } catch (error) {
//       // set a new state conserving the previous state + the new error
//       this.setState({ errors: [...this.state.errors, error] });
//     }
//   };

//   private _getErrors() {
//     return this.state.errors.length > 0 ? (
//       <div style={{ color: "orangered" }}>
//         <div>Errors:</div>
//         {this.state.errors.map((item, idx) => (
//           <div key={idx}>{JSON.stringify(item)}</div>
//         ))}
//       </div>
//     ) : null;
//   }
// }

// import * as React from "react";

// import React, { useEffect, useState } from "react";
// import styles from "./Demo.module.scss";

// import { IFile, IResponseItem } from "./interface";
// import { SPFI } from "@pnp/sp";
// import "@pnp/sp/webs";
// import "@pnp/sp/lists";
// import "@pnp/sp/items";

// export interface IAsyncAwaitPnPJsProps {
//   description: string;
//   sp: SPFI;
// }

// const Demo: React.FC<IAsyncAwaitPnPJsProps> = ({ description, sp }) => {
//   const [items, setItems] = useState<IFile[]>([]);
//   const [errors, setErrors] = useState<any[]>([]);

//   useEffect(() => {
//     const fetchData = async (libraryName: string): Promise<void> => {
//       try {
//         const response: IResponseItem[] = await sp.web.lists
//           .getByTitle(libraryName)
//           .items.select("Title", "FileLeafRef", "File/Length")
//           .expand("File/Length")();

//         const fetchedItems: IFile[] = response.map((item: IResponseItem) => ({
//           Id: item.Id,
//           Title: item.Title,
//           Size: item.File.Length,
//           Name: item.FileLeafRef
//         }));

//         setItems(fetchedItems);
//       } catch (error) {
//         setErrors(prevErrors => [...prevErrors, error]);
//       }
//     };

//     void fetchData("Documents");
//   }, [sp]);

//   const totalDocs: number = items.length > 0
//     ? items.reduce<number>((acc, item) => acc + Number(item.Size), 0)
//     : 0;

//   const renderErrors = () => (
//     errors.length > 0 && (
//       <div style={{ color: "orangered" }}>
//         <div>Errors:</div>
//         {errors.map((error, idx) => (
//           <div key={idx}>{JSON.stringify(error)}</div>
//         ))}
//       </div>
//     )
//   );

//   return (
//     <div className={styles.demo}>
//       <div className={`ms-Grid-row ms-bgColor-themeDark ms-fontColor-white ${styles.row}`}>
//         <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
//           <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
//           {renderErrors()}
//           <p className="ms-font-l ms-fontColor-white">List of documents:</p>
//           <div>
//             <div className={styles.row}>
//               <div className={styles.left}>Name</div>
//               <div className={styles.right}>Size (KB)</div>
//               <div className={`${styles.clear} ${styles.header}`} />
//             </div>
//             {items.map((item, idx) => (
//               <div key={idx} className={styles.row}>
//                 <div className={styles.left}>{item.Name}</div>
//                 <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
//                 <div className={styles.clear} />
//               </div>
//             ))}
//             <div className={styles.row}>
//               <div className={`${styles.clear} ${styles.header}`} />
//               <div className={styles.left}>Total: </div>
//               <div className={styles.right}>{(totalDocs / 1024).toFixed(2)}</div>
//               <div className={styles.clear} />
//             </div>
//           </div>
//         </div>
//       </div>
//     </div>
//   );
// };

// export default Demo;