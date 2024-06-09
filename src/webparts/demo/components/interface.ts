export interface IFile {
  Id: number;
  Title: string;
  Name: string;
    Size: number;
   //more metadata
  CreatedDate: Date;
  ModifiedDate: Date;
  Author: string;
  Editor: string;
}

export interface IResponseFile {
    Length: number;
    //more metadata
  TimeCreated: string;
  TimeLastModified: string;
  AuthorId: number;
  EditorId: number;
}

export interface IResponseItem {
  Id: number;
  File: IResponseFile;
  FileLeafRef: string;
   Title: string;
 //more metadata
  Created: string;
  Modified: string;
  AuthorId: number;
  EditorId: number;
}
