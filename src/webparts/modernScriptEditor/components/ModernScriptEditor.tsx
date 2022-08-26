import * as React from 'react';
import styles from './ModernScriptEditor.module.scss';
import { IModernScriptEditorProps } from './IModernScriptEditorProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { DisplayMode } from '@microsoft/sp-core-library';
import { Logger, LogLevel } from "@pnp/logging";

export default class ModernScriptEditor extends React.Component<IModernScriptEditorProps, any> {  
  private LOG_SOURCE = "ModernScriptEditor";
  
  constructor(props: IModernScriptEditorProps, state: any){
    super(props);

    this._showDialog = this._showDialog.bind(this);
    this.state = {      
      canView: false
    };
  }

  private _showDialog() {
    this.props.propPaneHandle.open();
  }

  private _checkUserCanViewWebpart = async (): Promise<void> => {
    try {
      Logger.write(`${this.LOG_SOURCE} (_checkUserCanViewWebpart)`, LogLevel.Verbose);

      let isPart_userGroups = false;
      if(this.props.userGroupInfos && this.props.audienceGroups && this.props.audienceGroups.length > 0) {
        this.props.userGroupInfos.forEach(userGroupInfo => {        
          this.props.audienceGroups.forEach(groupID => {
              if(groupID.id == userGroupInfo.Id.toString()) {
                isPart_userGroups = true;
              } else { }
          });
        });
      }

      let isPart_AZGroups = false;
      if(this.props.user_AZGroupInfos && this.props.audience_AZGroups && this.props.audience_AZGroups.length > 0) {
        this.props.user_AZGroupInfos.forEach(user_AZGroupInfo => {        
          // Logger.write(`(user_AZGroupInfo) - ${LogLevel.Verbose}\n${user_AZGroupInfo.displayName} ${user_AZGroupInfo.id}`, LogLevel.Verbose);
          this.props.audience_AZGroups.forEach(azGroup => {
            // Logger.write(`(azGroup) - ${LogLevel.Verbose}\n${azGroup.fullName} ${azGroup.id}`, LogLevel.Verbose);
              if(azGroup.fullName == user_AZGroupInfo.displayName) {
                isPart_AZGroups = true;
              } else {              
                Logger.write(`(user_AZGroupInfo) - Not Matched ${azGroup.fullName} ${azGroup.id} ${user_AZGroupInfo.displayName} ${user_AZGroupInfo.id}`, LogLevel.Verbose);
              }
          });
        });
      }

      //when a user is part of any selected groups
      if(isPart_userGroups || isPart_AZGroups) {
        this.setState({ fileContent: this.props.fileContent, script: this.props.script, canView: true });
      }

    } catch (err) {
      Logger.write(`${this.LOG_SOURCE} (_checkUserCanViewWebpart) - ${LogLevel.Warning}\n${JSON.stringify(err)}`, LogLevel.Warning);
    }
  }

  public componentDidMount(): void {
    
    if(this.props.fileContent || this.props.script) {
      // empty groups
      if (!this.props.audienceGroups && !this.props.audience_AZGroups) {
        this.setState({ fileContent: this.props.fileContent, script: this.props.script, canView: true });
      } //SP group deselected and AAD group not selected
      else if((this.props.audienceGroups && this.props.audienceGroups.length == 0 ) && (!this.props.audience_AZGroups)) {
        this.setState({ fileContent: this.props.fileContent, script: this.props.script, canView: true });
      } //SP group not selected and AAD group deselected
      else if(!this.props.audienceGroups && (this.props.audience_AZGroups && this.props.audience_AZGroups.length ==0)) {
        this.setState({ fileContent: this.props.fileContent, script: this.props.script, canView: true });
      } //both SP group and AAD group deselected
      else if((this.props.audienceGroups && this.props.audienceGroups.length == 0 ) && (this.props.audience_AZGroups && this.props.audience_AZGroups.length == 0)) {
          this.setState({ fileContent: this.props.fileContent, script: this.props.script, canView: true });
      } // SP group or/and AAD group selected
      else
      {
        this._checkUserCanViewWebpart();
      }
    }
  }

  public render(): React.ReactElement<IModernScriptEditorProps> {
    const {      
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      displayMode,
      title,      
      fileContent,
      script,
      userGroupInfos,
      user_AZGroupInfos,
      audienceGroups,
      audience_AZGroups,
    } = this.props;
    
    return (
      <div className='ms-Fabric'>
        {(!script && !fileContent) && <Placeholder 
            iconName='FileHTML'
            iconText={title}
            description='Please configure the web part'
            buttonLabel='Edit markup'
            onConfigure={this._showDialog} />}        
        {(displayMode==DisplayMode.Edit) && (fileContent || script) &&
          (<div className={ styles.modernScriptEditor }>
            <span className={ styles.title }>{escape(title)}</span>
          </div>)}
        {this.state.canView && fileContent ?
          (<div className={styles.contentLink}>
            <span dangerouslySetInnerHTML={{ __html: this.state.fileContent }}></span>
          </div>) : '' }
        {this.state.canView && script ?
          (<div className={ styles.modernScriptEditor }>
            <span dangerouslySetInnerHTML={{ __html: this.state.script }}></span>
          </div>) : ''
        }
      </div>
    );
  }
}
