/* eslint-disable @typescript-eslint/explicit-function-return-type */
/* eslint-disable @typescript-eslint/no-explicit-any */
/* eslint-disable @typescript-eslint/no-use-before-define */
/* eslint-disable @typescript-eslint/no-use-before-define */
import React, { useEffect, useState, useCallback, useMemo } from "react";
import styles from "./Demo.module.scss";
import { IFile, IResponseItem } from "./interface";
import { SPFI, spfi, SPFx, RequestDigest } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/batching";
import "@pnp/sp/site-users/web";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import "@pnp/sp/sharing";
import { Logger, LogLevel } from "@pnp/logging";
import { Caching } from "@pnp/queryable";
import Bottleneck from 'bottleneck';
import { User } from 'lucide-react';
export interface IDemoProps {
  description: string;
  sp: SPFI;
  context: any;
}

const Demo: React.FC<IDemoProps> = ({ context }) => {
  const [items, setItems] = useState<IFile[]>([]);
  const [errors, setErrors] = useState<string[]>([]);
  const [userEmail, setUserEmail] = useState<string>("");
  const [newTitles, setNewTitles] = useState<{ [key: number]: string }>({});
  const LOG_SOURCE = "Demo";
  const LIBRARY_NAME = "Documents";

  const limiter = useMemo(() => new Bottleneck({
    minTime: 200
  }), []);

  const spInstance = useMemo(() => spfi()
    .using(SPFx(context))
    .using(RequestDigest())
    .using(Caching({ store: "session", keyFactory: (url) => `${url}:${userEmail}` })), [context, userEmail]);
  console.log(spInstance);

  const getCurrentUserEmail = useCallback(async (): Promise<void> => {
    try {
      const user = await spInstance.web.currentUser();
      setUserEmail(user.Email);
      console.log(`Current user email retrieved: ${user.Email}`);
      Logger.write(`Current user email: ${user.Email}`, LogLevel.Info);
    } catch (error) {
      Logger.write(`Error getting current user email: ${JSON.stringify(error)}`, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  }, [spInstance]);

  const handleThrottling = async (fn: () => Promise<any>, retries = 3): Promise<any> => {
    try {
      return await fn();
    } catch (error) {
      if (retries > 0 && (error.status === 429 || error.status === 503)) {
        const retryAfter = error.headers['Retry-After'] || 1;
        await new Promise(resolve => setTimeout(resolve, retryAfter * 1000));
        return handleThrottling(fn, retries - 1);
      } else {
        throw error;
      }
    }
  };

  const readAllFilesSize = useCallback(async (libraryName: string): Promise<void> => {
    try {
      const response: IResponseItem[] = await handleThrottling(() =>
        limiter.schedule(() =>
          spInstance.web.lists
            .getByTitle(libraryName)
            .items.select("Id", "Title", "FileLeafRef", "File/Length")
            .expand("File/Length")()
        )
      );

      const files: IFile[] = response.map((item: IResponseItem) => ({
        Id: item.Id,
        Title: item.Title,
        Size: item.File.Length,
        Name: item.FileLeafRef,
      }));

      console.log(`Files retrieved: ${JSON.stringify(files)}`);
      setItems(files);
    } catch (error) {
      Logger.write(`${LOG_SOURCE} (readAllFilesSize) - ${JSON.stringify(error)} - `, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  }, [spInstance, limiter]);

  useEffect(() => {
    const fetchData = async () => {
      await getCurrentUserEmail();
      await readAllFilesSize(LIBRARY_NAME);
    };

    fetchData().catch(error => {
      Logger.write(`Error initializing data: ${error}`, LogLevel.Error);
    });
  }, [getCurrentUserEmail, readAllFilesSize, LIBRARY_NAME]);

  const batchUpdateItemTitles = useCallback(async (): Promise<void> => {
    try {
      const [batchedSP, execute] = spInstance.batched();
      const list = batchedSP.web.lists.getByTitle(LIBRARY_NAME);

      for (const item of items) {
        list.items.getById(item.Id).update({ Title: `${item.Name}-updated` }).catch(async (error) => {
          if (error.message.includes('is locked for shared use')) {
            const lockedByUser = await getLockedByUser(item.Name);
            const errorMessage = `File is locked: ${item.Name} by ${lockedByUser}`;
            setErrors(prevErrors => [...prevErrors, errorMessage]);
          } else {
            throw error;
          }
        });
      }

      await execute();
      console.log(`Batch update executed`);
      await readAllFilesSize(LIBRARY_NAME);
    } catch (error) {
      Logger.write(`Error batch updating item titles: ${JSON.stringify(error)}`, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
    }
  }, [spInstance, items, readAllFilesSize]);

  const updateItemTitle = useCallback(async (itemId: number, newTitle: string): Promise<void> => {
    try {
      await spInstance.web.lists.getByTitle(LIBRARY_NAME).items.getById(itemId).update({ Title: newTitle });
      console.log(`Item ${itemId} updated to title ${newTitle}`);
      await readAllFilesSize(LIBRARY_NAME);
    } catch (error) {
      Logger.write(`Error updating item title: ${JSON.stringify(error)}`, LogLevel.Error);
      setErrors(prevErrors => [...prevErrors, error.message]);
      console.error(`Error in updateItemTitle for item ${itemId}:`, error);
    }
  }, [spInstance, readAllFilesSize]);

  const handleTitleChange = (itemId: number, newTitle: string) => {
    setNewTitles(prevTitles => ({
      ...prevTitles,
      [itemId]: newTitle
    }));
  };

  const getLockedByUser = useCallback(async (fileName: string): Promise<string> => {
    try {
      const file = spInstance.web.getFolderByServerRelativePath(LIBRARY_NAME).files.getByUrl(fileName);
      const user = await file.getLockedByUser();
      console.log(`Locked by user retrieved: ${user?.Email}`);
      return user?.Email || 'Unknown user';
    } catch (error) {
      Logger.write(`Error getting locked by user for file ${fileName}: ${JSON.stringify(error)}`, LogLevel.Error);
      console.error("Error in getLockedByUser:", error);
      return 'Unknown user';
    }
  }, [spInstance, LIBRARY_NAME]);

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

  const totalDocs: number = items.reduce((acc: number, item: IFile) => acc + Number(item.Size), 0);

  return (
       <div className={styles.demo}>
      <div className={"ms-Grid-row ms-bgColor-themeDark ms-fontColor-white " + styles.row}>
        <div className="ms-Grid-col ms-u-lg10 ms-u-xl8 ms-u-xlPush2 ms-u-lgPush1">
          <span className="ms-font-xl ms-fontColor-white">Welcome to Yehfedra SharePoint PnP/sp Demo</span>
          {getErrors()}

          <p className="ms-font-l ms-fontColor-white">Current User Email: <User color="blue" size={20} /> {userEmail}</p>
          <p className="ms-font-l ms-fontColor-white">List of documents:</p>
          <div>
            <div className={styles.row}>
              <div className={styles.left}>Name</div>
              <div className={styles.right}>Size (KB)</div>
              <div className={styles.left}>Locked By</div>
              <div className={`${styles.clear} ${styles.header}`} />
            </div>
            {items.map((item, idx) => (
              <div key={idx} className={styles.row}>
                <div className={styles.left}>{item.Name}</div>
                <div className={styles.left}>{item.Title}</div>
                <div className={styles.left}>
                  <input
                    type="text"
                    value={newTitles[item.Id] !== undefined ? newTitles[item.Id] : item.Title}
                    onChange={(e) => handleTitleChange(item.Id, e.target.value)}
                  />
                </div>
                <div className={styles.right}>{(item.Size / 1024).toFixed(2)}</div>
                {/* <div className={styles.left}>{item.LockedUser}</div> */}
                <button onClick={() => updateItemTitle(item.Id, newTitles[item.Id] !== undefined ? newTitles[item.Id] : item.Title)}>
                  Update Title
                </button>
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
          <button onClick={() => batchUpdateItemTitles()}>Update Item Titles</button>
        </div>
      </div>
    </div>
  );
};

export default Demo;
