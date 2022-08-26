import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, DisplayMode } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { 
  PropertyFieldPeoplePicker, 
  PrincipalType, 
  IPropertyFieldGroupOrPerson } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'ModernScriptEditorWebPartStrings';
import ModernScriptEditor from './components/ModernScriptEditor';
import { IModernScriptEditorProps } from './components/IModernScriptEditorProps';
import { getSP, CustomListener, getHttpClient } from './pnpjsConfig';
import { Logger, LogLevel, ConsoleListener } from "@pnp/logging";
import UserGroupLookup from '../../services/UserGroupLookup';
import { ISiteGroupInfo } from '@pnp/sp/site-groups/types';

export interface IModernScriptEditorWebPartProps {
  title:string;
  script: string;
  removePadding:boolean;
  contentLink: string;
  fileContent: string;
  targetAudienceGroups: IPropertyFieldGroupOrPerson[];
  targetAudience_AZGroups: IPropertyFieldGroupOrPerson[];
  teamsContext: boolean;
}

export default class ModernScriptEditorWebPart extends BaseClientSideWebPart<IModernScriptEditorWebPartProps> {
  private LOG_SOURCE = "ModernScriptEditorWebPart";
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private _uniqueId;
  private _scriptUniqueAttr: string = "cibctoday_modernscript";
  public _propertyPaneHelper;

  private _userGroups: UserGroupLookup;
  private _userGroupInfos: ISiteGroupInfo[];
  private _user_AZGroupInfos: any[];
  
  constructor() {
    super();
    this.targetAudienceGroupsUpdated = this.targetAudienceGroupsUpdated.bind(this);
    this.targetAudience_AZGroupsUpdated = this.targetAudience_AZGroupsUpdated.bind(this);
    this.scriptUpdated = this.scriptUpdated.bind(this);    
  }  

  protected async onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();
    
    await super.onInit();

    Logger.activeLogLevel = LogLevel.Warning;
    Logger.subscribe(new CustomListener());    

    // Initialize our _sp object that we can then use in other packages without having to pass around the context.    
    getSP(this.context);

    //Initialize _HttpClient used in other packages without having to pass around the context.
    getHttpClient(this.context);
    
    // Initialize the groups that the current user is belong to
    this._loadCurrentUserGroups();

    // Initialize the AAD groups that the current user is belong to
    this._loadCurrentUserAZGroups();
  }

  // Request to load the SP groups where the current user belongs to
  private _loadCurrentUserGroups = async (): Promise<void> => {
    try {

      Logger.write(`${this.LOG_SOURCE} (_loadCurrentUserGroups---)`, LogLevel.Verbose);
      
      if(!this._userGroups) {
        this._userGroups = new UserGroupLookup();
      }

      if(!this._userGroupInfos) {
          this._userGroupInfos = await this._userGroups.getCurrentUserGroups();
      }

      Logger.write(`${this.LOG_SOURCE} (_loadCurrentUserGroups done)`, LogLevel.Verbose);
      this.render();
      
    } catch (err) {
      Logger.write (`${this.LOG_SOURCE} (_loadCurrentUserGroups) - ${LogLevel.Error}\n${err.message}`, LogLevel.Error);      
    }
  }

  // Request to load the AAD groups where the current user belongs to
  private _loadCurrentUserAZGroups = async (): Promise<void> => {
    try {
      
      Logger.write(`${this.LOG_SOURCE} (_loadCurrentUserAZGroups---)`, LogLevel.Verbose);
      
      if(!this._userGroups) {
        this._userGroups = new UserGroupLookup();
      }

      if(!this._user_AZGroupInfos) {
        this._user_AZGroupInfos = await this._userGroups.getCurrentUser_AZGroups(this.context.pageContext.user.loginName, this.context.pageContext.site.absoluteUrl.toLowerCase());        
      }
      
      this._user_AZGroupInfos.forEach(item => {
        if(item.custom && item.custom.length > 0) {
          if(item.custom.indexOf('404') > -1) {
            Logger.write (`${this.LOG_SOURCE} (_loadCurrentUserAZGroups) - ${LogLevel.Warning}\n${item.custom}`, LogLevel.Warning);            
          }
        }
      });

      Logger.write(`${this.LOG_SOURCE} (_loadCurrentUserAZGroups done)`, LogLevel.Verbose);
      this.render();

    } catch (err) {      
      Logger.write (`${this.LOG_SOURCE} (_loadCurrentUserAZGroups) - ${LogLevel.Error}\n${err.message}`, LogLevel.Error);
    }
  }

  public targetAudienceGroupsUpdated(_property: string, _oldVals: IPropertyFieldGroupOrPerson[], newVals: IPropertyFieldGroupOrPerson[]) {    
    if(!this._userGroupInfos) { this._loadCurrentUserGroups() }
    // newVals.forEach(newVal => {Logger.write("    " + newVal.id + ' ' + newVal.fullName)});
    //Logger.write(`${this.LOG_SOURCE} (targetAudienceGroupsUpdate) - ${LogLevel.Verbose}`);
  }

  public targetAudience_AZGroupsUpdated(propertyPath: string, oldValue: any, newValue: any) {
    if(!this._user_AZGroupInfos) { this._loadCurrentUserAZGroups() }
    // this.properties.targetAudience_AZGroups = newValue;
    //Logger.write(`${this.LOG_SOURCE} (targetAudience_AZGroupsUpdated) - ${LogLevel.Verbose}`);
  }

  public scriptUpdated(_property: string, _oldVal: string, newVal: string) : void {
    this.properties.script = newVal;
    this._propertyPaneHelper.initialValue = newVal;
    //Logger.write(`${this.LOG_SOURCE} (scriptUpdated) - ${LogLevel.Verbose} \n`);
  }

  //Check if a string is empty or not
  private emptyOrWhiteSpace(str: string) {
    return str == null || str.replace(/\s/g, "").length < 1;
  }
  
  //Check if a file type is html or txt
  private isValidExt(str: string) {
    let url = str.split(".");
    let extension = url[url.length - 1].toLowerCase();
    let res = extension.match(/html|txt/);
    return (res !== null);
  }

  // Valide the field of contentLink
  private async validateContentLink(path: string) : Promise<string> {
    try {     

      if (this.emptyOrWhiteSpace(path)) {
        this.properties.fileContent = '';
        return '';
      }
      
      //Check file extension exist to go forward to fetch
      if(!this.isValidExt(path)){
        return "Please type a valid html or txt file";
      }

      //Fetch the file content when the contentLink field is not empty, and has valid file type.
      const headers = new Headers();
      headers.set("Access-Control-Allow-Origin", "*");
      const init: RequestInit = {
        method: "GET",
        headers,
      };

      const response = await fetch(path, init);

      if (response.ok) {
        this.properties.fileContent = await response.text();          
        Logger.write(`${this.LOG_SOURCE} (validateContentLink) - ${LogLevel.Verbose} \n ${this.properties.contentLink} \n ${this.properties.fileContent}`, LogLevel.Verbose);
        return "";
      } else if (response.status === 404) {
        return `The file you provided doesn't exist`;
      }
      else {
        Logger.write(`${this.LOG_SOURCE} (validateContentLink) - ${LogLevel.Info} \n $${response.status} ${response.statusText}`);
        return `Please try again. ${response.statusText}`;
      }
    } catch (error) {
      Logger.write(`${this.LOG_SOURCE} (validateContentLink) - ${LogLevel.Error} \n ${error}`, LogLevel.Error);
      return error.message;
    }
  }
  
  public render(): void {
    Logger.write(`${this.LOG_SOURCE} (render)`, LogLevel.Verbose);
    this._uniqueId = this.context.instanceId;

    if(this.displayMode == DisplayMode.Edit) {
      this.renderEditor();      
    } else {
      if(this.properties.removePadding){
        let element = this.domElement.parentElement;
        // Chekc up to 5 levels for paddding and exit once found
        for(let i =0; i < 5; i++){
          const style = window.getComputedStyle(element);
          const hasPadding = style.paddingTop ! == "0px";
          if(hasPadding){
            element.style.paddingTop = "0px";
            element.style.paddingBottom = "0px";
            element.style.marginTop = "0px";
            element.style.marginBottom = "0px";
          }
          element = element.parentElement;
        }
      }
      
      //ReactDom.unmountComponentAtNode(this.domElement);
      const element: React.ReactElement<IModernScriptEditorProps> = React.createElement(
        ModernScriptEditor,
        {
          isDarkTheme: this._isDarkTheme,
          environmentMessage: this._environmentMessage,
          hasTeamsContext: !!this.context.sdks.microsoftTeams,
          userDisplayName: this.context.pageContext.user.displayName,
          displayMode: this.displayMode,
          title: this.properties.title,
          script: this.properties.script,
          pageContext: this.context.pageContext,
          userGroupInfos: this._userGroupInfos,
          user_AZGroupInfos: this._user_AZGroupInfos,
          audienceGroups: this.properties.targetAudienceGroups,
          audience_AZGroups: this.properties.targetAudience_AZGroups,
          contentLink: this.properties.contentLink,
          fileContent: this.properties.fileContent,
          propPaneHandle: this.context.propertyPane,
          key: new Date().getTime()
        }
      );
      ReactDom.render(element, this.domElement);
      
      this.executeScript(this.domElement);
    }
  }

  private async renderEditor() {
    // Dynamically load the editor pane to reduce overall bundle size
    const editorPopUp = await import('./components/ModernScriptEditor');
     
    const element: React.ReactElement<IModernScriptEditorProps> = React.createElement(
      editorPopUp.default,
      {
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        displayMode: this.displayMode,
        title: this.properties.title,
        script: this.properties.script,
        pageContext: this.context.pageContext,        
        userGroupInfos: this._userGroupInfos,
        user_AZGroupInfos: this._user_AZGroupInfos,
        audienceGroups: this.properties.targetAudienceGroups,
        audience_AZGroups: this.properties.targetAudience_AZGroups,
        contentLink: this.properties.contentLink,
        fileContent: this.properties.fileContent,
        propPaneHandle: this.context.propertyPane,
        key: new Date().getTime()
      }
    );
    ReactDom.render(element, this.domElement);
  }

  // Finds and executes scripts in a newly added elements body as innerHTML does not run scripts
  // Argument HTMLElment is an element in the dom
  private async executeScript(element: HTMLElement){

    // clean up added script tags in case of smart reload    
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;    
    let scriptTags = headTag.getElementsByTagName("script");
    for(let i = 0; i < scriptTags.length; i++){
      const scriptTag = scriptTags[i];
      if(scriptTag.hasAttribute(this._scriptUniqueAttr) && scriptTag.attributes[this._scriptUniqueAttr].value == this._uniqueId)
      {
        headTag.removeChild(scriptTag);
        Logger.write(`${this.LOG_SOURCE} (executeScript) scriptTag removed`, LogLevel.Verbose);
      }
    }

    if(this.properties.teamsContext && !window["_teamsContexInfo"]) {
      window["_teamsContexInfo"] = this.context.sdks.microsoftTeams.context;
    }

    // Define global name to tack scripts on in case script to be loaded is not AMD/UMD
    (<any>window).ScriptGlobal = {};

    // maind sectin of function
    const scripts = [];
    const children_nodes = element.getElementsByTagName("script");

    for(let i = 0; children_nodes[i]; i++){
      const child: any = children_nodes[i];
      if(!child.type || child.type.toLowerCase() === "text/javascript") {
        scripts.push(child);
      }
    }

    const urls = [];
    const onLoads = [];
    for(let i=0; scripts[i]; i++){
      const scriptTag = scripts[i];
      if(scriptTag.src && scriptTag.src.length > 0){
        urls.push(scriptTag.src);
      }

      if(scriptTag.onLoad && scriptTag.onLoad.length > 0){
        onLoads.push(scriptTag.onLoad);
      }
    }

    let oldamd = null;
    if(window["define"] && window["define"].amd) {
      oldamd = window["define"].amd;
      window["define"].amd = null;
    }

    for(let i = 0; i < urls.length; i++){
      try{
        let scriptUrl = urls[i];
        //Add unique param to force load on each run to overcome smart navigation in the browser as needed
        const prefix = scriptUrl.indexOf('?') === -1 ? '?' : '&';
        scriptUrl += prefix + this._scriptUniqueAttr + new Date().getTime();
        await SPComponentLoader.loadScript(scriptUrl, {globalExportsName: "ScriptGlobal"});
      } catch (error) {
        if(console.error){
          console.error(error);
        }
      }      
    }

    if(oldamd) {
      window["define"].amd = oldamd;
    }

    for(let i =0; scripts[i]; i++) {
      const scriptTag = scripts[i];
      if(scriptTag.parentNode) {
        scriptTag.parentNode.removeChild(scriptTag);
      }
      this.evalScript(scripts[i]);
    }
    //execute any onload that has added
    for(let i =0; onLoads[i]; i++) {
      onLoads[i]();
    }

    Logger.write(`${this.LOG_SOURCE} (executeScript done) - ${LogLevel.Verbose} \n ${this.domElement.innerHTML}`, LogLevel.Verbose);
  }
  
  private evalScript(element) {
    const data = (element.text || element.textContent || element.innerHTML || "");    
    const headTag = document.getElementsByTagName("head")[0] || document.documentElement;
    const scriptTag = document.createElement("script");

    for(let i = 0; i < element.attributes.length; i++){
      const attr = element.attributes[i];

      // Copies all attributes in case of loaded script relies on the tag attributes
      if(attr.name.toLowerCase() === "onload")
        continue;
      scriptTag.setAttribute(attr.name, attr.value);
    }

    // set a bogus type to avoid browser loading the script, as it's loaded with SPComponentloader
    scriptTag.type = (scriptTag.src && scriptTag.src.length) > 0 ? this._scriptUniqueAttr : "text/javascript";
    // Ensure proper setting and adding id used in cleanup on reload
    scriptTag.setAttribute(this._scriptUniqueAttr, this._uniqueId);

    try {
      // does not work on IE...
      scriptTag.appendChild(document.createTextNode(data));
    } catch(e) {
      //IE has a bit weird script nodes
      scriptTag.text = data;
    }

    headTag.insertBefore(scriptTag, headTag.firstChild);    
    Logger.write(`${this.LOG_SOURCE} (evalScript) scriptTag added`, LogLevel.Verbose);
  }

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {    
    return Version.parse('1.2');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    let webPartOptions: IPropertyPaneField<any>[] = 
    [
      PropertyPaneTextField("title", {
        label: "Title in Edit mode",
        value: this.properties.title
      })
      ,
      PropertyPaneToggle("removePadding", {
        label: "Remove padding of web part container",
        checked: this.properties.removePadding,
        onText: "Remove padding",
        offText: "Keep padding"
      }),
      PropertyFieldPeoplePicker('targetAudienceGroups', {
        label: 'Target Audience - SP Groups',
        initialData: this.properties.targetAudienceGroups,
        allowDuplicate: false,
        principalType: [PrincipalType.SharePoint],
        onPropertyChange: this.targetAudienceGroupsUpdated,
        context: this.context as any,
        properties: this.properties,
        onGetErrorMessage: null,
        deferredValidationTime: 0,
        key: 'peopleFieldId'
      }),
      PropertyFieldPeoplePicker('targetAudience_AZGroups', {
        label: 'Target Audience - AAD Groups',
        initialData: this.properties.targetAudience_AZGroups,
        allowDuplicate: false,
        principalType: [PrincipalType.Security],
        onPropertyChange: this.targetAudience_AZGroupsUpdated,
        context: this.context as any,
        properties: this.properties,
        onGetErrorMessage: null, //this._loadCurrentUserAZGroups.bind(this),
        deferredValidationTime: 0,
        key: 'people_AZ_FieldId'
      }),
      PropertyPaneTextField("contentLink", {
        label: "To link to a html/text file, type a URL",
        multiline: true,
        resizable: true,
        onGetErrorMessage: this.validateContentLink.bind(this),
        placeholder: "",
      }),
      this._propertyPaneHelper
    ];

    if(this.context.sdks.microsoftTeams){
      let config = PropertyPaneToggle("teamsContext", {
        label: "Enable teams context as _teasmContextInfo",
        checked: this.properties.teamsContext,
        onText: "Enalbed",
        offText: "Disabled"
      });
      webPartOptions.push(config);
    }
    
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: webPartOptions
            }
          ]
        }
      ]
    };
  }

    // // Diable automatic property changes to take place (Apply button will have to be clicked to apply changes)
    // protected get disableReactivePropertyChanges(): boolean {
    //   return true;
    // }
 
  protected async loadPropertyPaneResources(): Promise<void> {
    const editorProp = await import('@pnp/spfx-property-controls/lib/PropertyFieldCodeEditor');

    this._propertyPaneHelper = editorProp.PropertyFieldCodeEditor('scriptCode', {
      label: "Edit HTML Code",
      panelTitle: 'Edit HTML Code',
      initialValue: this.properties.script,
      onPropertyChange: this.scriptUpdated,
      properties: this.properties,
      disabled: false,
      key: 'codeEditorFieldId',
      language: editorProp.PropertyFieldCodeEditorLanguages.HTML,
      options: {
        wrap: true,
        fontSize: 12
      }
    });
  }
}
