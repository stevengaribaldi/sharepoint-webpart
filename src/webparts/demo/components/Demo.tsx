

import React, { useEffect, useState } from "react";
import styles from "./Demo.module.scss";
import { IFile, IResponseItem } from "./interface";
import {
  SPFI,
  // spfi, RequestDigest
} from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/sharing";
import { Logger, LogLevel } from "@pnp/logging";
// import { SharingRole } from "@pnp/sp/sharing";

export interface IDemoProps {
  description: string;
  sp: SPFI;
}

const Demo: React.FC<IDemoProps> = ({ description, sp, }) => {
  const [items, setItems] = useState<IFile[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [userEmail, setUserEmail] = useState<string>("");
  const LOG_SOURCE = "Demo";
  const LIBRARY_NAME = "Documents";

  // const sp = spfi().using(SPFx(context));  // Initialize sp with SPFx context


  const getCurrentUserEmail = async (): Promise<void> => {
    try {
      const user = await sp.web.currentUser();
      setUserEmail(user.Email);
      Logger.write(`Current user email: ${user.Email}`, LogLevel.Info);
    } catch (error) {
      Logger.write(`Error getting current user email: ${JSON.stringify(error)}`, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  };

  const readAllFilesSize = async (libraryName: string): Promise<void> => {
    try {
      const response: IResponseItem[] = await sp.web.lists
        .getByTitle(libraryName)
        .items.select("Id", "Title", "FileLeafRef", "File/Length")
        .expand("File/Length")();

      const files: IFile[] = response.map((item: IResponseItem) => ({
        Id: item.Id,
        Title: item.Title,
        Size: item.File.Length,
        Name: item.FileLeafRef,
      }));

      setItems(files);
    } catch (error) {
      Logger.write(`${LOG_SOURCE} (readAllFilesSize) - ${JSON.stringify(error)} - `, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  };

  const batchUpdateItemTitles = async (): Promise<void> => {
    try {
      const [batchedSP, execute] = sp.batched();
      const list = batchedSP.web.lists.getByTitle(LIBRARY_NAME);

      for (const item of items) {
        list.items.getById(item.Id).update({ Title: `${item.Name}-Updatbcvvvvvved` }).catch(async (error) => {
          if (error.message.includes('is locked for shared use')) {
            // eslint-disable-next-line @typescript-eslint/no-use-before-define
            const lockedByUser = await getLockedByUser(item.Name);
            const errorMessage = `File is locked: ${item.Name} by ${lockedByUser}`;
            setErrors(prevErrors => [...prevErrors, errorMessage]);
          } else {
            throw error;
          }
        });
      }

      await execute();
      await readAllFilesSize(LIBRARY_NAME);
    } catch (error) {
      Logger.write(`Error batch updating item titles: ${JSON.stringify(error)}`, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  };

  const getLockedByUser = async (fileName: string): Promise<string> => {
    try {
      const file = sp.web.getFolderByServerRelativePath(LIBRARY_NAME).files.getByUrl(fileName);
      const user = await file.getLockedByUser();
      return user?.Email || 'Unknown user';
    } catch (error) {
      Logger.write(`Error getting locked by user for file ${fileName}: ${JSON.stringify(error)}`, LogLevel.Error);
      return 'Unknown user';
    }
  };

  const getErrors = (): JSX.Element | null => {
    return errors.length > 0 ? (
      <div style={{ color: "orangered" }}>
        <div>Errors:</div>
        {errors.map((item, idx) => (
          <div key={idx}>{item}</div>
        ))}
      </div>
    ) : null;
  };

  useEffect(() => {
    const fetchData = async () => {
      await getCurrentUserEmail();
      await readAllFilesSize(LIBRARY_NAME);
    };

    fetchData().catch(error => {
      Logger.write(`Error initializing data: ${error}`, LogLevel.Error);
    });
  }, [getCurrentUserEmail, readAllFilesSize, LIBRARY_NAME]);
  const totalDocs: number = items.reduce((acc: number, item: IFile) => acc + Number(item.Size), 0);

  return (
    <div className={styles.demo}>
      <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
        <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span className="ms-font-xl ms-fontColor-white">Welcome to SharePoint Async Await SP PnP JS Demo!</span>
          {getErrors()}
          <p className="ms-font-l ms-fontColor-white">Current User Email: {userEmail}</p>
          <p className="ms-font-l ms-fontColor-white">List of documents:</p>
          <div>
            <div className={styles.row}>
              <div className={styles.left}>Name</div>
              <div className={styles.right}>Size (KB)</div>
              <div className={`${styles.clear} ${styles.header}`} />
            </div>
            {items.map((item, idx) => (
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
          <button onClick={() => readAllFilesSize(LIBRARY_NAME)}>Refresh Files</button>
          <button onClick={() => batchUpdateItemTitles()}>Batch Update Item Titles</button>
        </div>
      </div>
    </div>
  );
};

// const sp = spfi().using(RequestDigest());

// async function grantAccess(resourceUrl: string, userEmail: string, role: SharingRole = SharingRole.View, isFolder: boolean = false) {
//   try {
//     if (isFolder) {
//       const result = await sp.web.getFolderByServerRelativePath(resourceUrl).shareWith(userEmail, role, true);
//       console.log(`Folder shared successfully: ${JSON.stringify(result, null, 2)}`);
//     } else {
//       const result = await sp.web.getFileByServerRelativePath(resourceUrl).shareWith(userEmail, role);
//       console.log(`File shared successfully: ${JSON.stringify(result, null, 2)}`);
//     }
//   } catch (error) {
//     console.error("Error sharing resource: ", error);
//   }
// }

// Usage
// const folderUrl = "/sites/dev/Shared Documents/folder1";
// const fileUrl = "/sites/dev/Shared Documents/test.txt";
// const userEmail = "i:0#.f|membership|user@site.com";

// Share a folder with edit permissions
// grantAccess(folderUrl, userEmail, SharingRole.Edit, true);

// // Share a file with view permissions
// grantAccess(fileUrl, userEmail, SharingRole.View, false);

export default Demo;
