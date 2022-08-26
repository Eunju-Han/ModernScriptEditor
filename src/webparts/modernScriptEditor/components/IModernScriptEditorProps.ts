import { IPropertyPaneAccessor } from "@microsoft/sp-webpart-base";
import { PageContext } from "@microsoft/sp-page-context";
import { DisplayMode } from '@microsoft/sp-core-library';
import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";
import { ISiteGroupInfo } from "@pnp/sp/site-groups/types";
import { IAZGroupInfo } from "./IAZGroupInfo";

export interface IModernScriptEditorProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;  
  displayMode: DisplayMode;
  title: string;
  script: string;
  pageContext: PageContext;  
  userGroupInfos: ISiteGroupInfo[];
  user_AZGroupInfos: IAZGroupInfo[];
  audienceGroups: IPropertyFieldGroupOrPerson[];
  audience_AZGroups: IPropertyFieldGroupOrPerson[];
  contentLink: string;
  fileContent: string;
  propPaneHandle: IPropertyPaneAccessor;
}
