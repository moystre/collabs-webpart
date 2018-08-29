import { ISPList } from "../../../../lib/webparts/collaborators/CollaboratorsWebPart";
import { ISPLists } from "../CollaboratorsWebPart";
import { IColumn } from "office-ui-fabric-react/lib/DetailsList";

export interface ICollaboratorsProps {
  description: string;
  ispLists: string[];
  columns: IColumn[];
}
