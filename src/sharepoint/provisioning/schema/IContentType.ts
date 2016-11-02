"use strict";
import {IContentTypeFieldRef} from "./IContentTypeFieldRef";

export interface IContentType {
  ID: SP.Guid;
  ParentId: SP.Guid;
  Name: string;
  Description?: string;
  Group: string;
  Hidden?: boolean;
  Sealed?: boolean;
  ReadOnly?: boolean;
  Overwrite?: boolean;
  NewFormUrl?: string;
  EditFormUrl?: string;
  DisplayFormUrl?: string;
  DocumentTemplate?: string;
  FieldRefs?: IContentTypeFieldRef[];
  DocumentSetTemplate?: string;
}
