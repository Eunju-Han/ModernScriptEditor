import { getSP, getHttpClient } from "../webparts/modernScriptEditor/pnpjsConfig"
import { SPFI, spfi } from "@pnp/sp";
import { Caching } from '@pnp/queryable';
import { Logger, LogLevel } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";
import { ISiteGroupInfo } from "@pnp/sp/site-groups/types";
import { HttpClient, AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { IAZGroupInfo } from "../webparts/modernScriptEditor/components/IAZGroupInfo";

export default class UserGroupLookup {

  private LOG_SOURCE = "ModernScriptEditorWebPart - UserGroupLookup";
    private _sp: SPFI;
    private _httpClient: HttpClient;

    private _userGroupInfos: ISiteGroupInfo[];    
    private _user_AZGroupInfo: IAZGroupInfo[];

    constructor() {
        this._sp = getSP();        
        this._httpClient = getHttpClient();
    }
    
    public async getCurrentUserGroups(): Promise<ISiteGroupInfo[]> {
        try {
            if(!this._userGroupInfos) {
                await this._retrieveCurrentUserGroups();
                return this._userGroupInfos;
            } else {
                return undefined;
            }            
        } catch (err) {
            Logger.write(`${this.LOG_SOURCE} (getCurrentUserGroups) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Verbose);             
            throw Error(`${err.message}`);
        }
    }
    
    public async getCurrentUser_AZGroups(upn: string, siteUrl: string): Promise<any[]> {
        try {
            if(!this._user_AZGroupInfo) {
                await this._retrieveCurrentUser_AZGroupsMemeberOf(upn, siteUrl);
                return this._user_AZGroupInfo;
            } else {
                return undefined;
            }            
        } catch (err) {
            Logger.write(`${this.LOG_SOURCE} (getCurrentUser_AZGroups) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Verbose);
            throw Error(`${err.message}`);
        }
    }

    private _retrieveCurrentUserGroups = async (): Promise<void> => {
        try {            
            // Creating a new sp object to include caching behavior. This way our original object is unchanged.
            const spCache = spfi(this._sp).using(Caching({ store: "session" }));
            Logger.write(`${this.LOG_SOURCE} (spCache) - ${LogLevel.Verbose} - ${spCache? true:false}`, LogLevel.Verbose);
            let response: ISiteGroupInfo[] = await spCache.web.currentUser.groups();

            Logger.write(`${this.LOG_SOURCE} _retrieveCurrentUserGroups - using spCache - succeeded - ${LogLevel.Info}`, LogLevel.Verbose);
            Logger.writeJSON(`${JSON.stringify(response)}`, LogLevel.Verbose);

            // Use map to convert ISiteGroupInfo[] into the internal object ISiteGroupInfo[]            
            this._userGroupInfos = response.map((item:ISiteGroupInfo) => {
                return {                    
                    Description: item.Description,
                    Id: item.Id,
                    Title: item.Title,
                    LoginName: item.LoginName,
                    OwnerTitle: item.OwnerTitle,
                    PrincipalType: item.PrincipalType,
                    IsHiddenInUI: item.IsHiddenInUI,
                    AllowMembersEditMembership: item.AllowMembersEditMembership,
                    AllowRequestToJoinLeave: item.AllowRequestToJoinLeave,
                    AutoAcceptRequestToJoinLeave: item.AutoAcceptRequestToJoinLeave,
                    RequestToJoinLeaveEmailSetting: item.RequestToJoinLeaveEmailSetting,
                    OnlyAllowMembersViewMembership: item.OnlyAllowMembersViewMembership
                }
            });            
            Logger.write(`${this.LOG_SOURCE} (_retrieveCurrentUserGroups done) - ${LogLevel.Verbose}\n}`, LogLevel.Verbose);

        } catch (err) {
            Logger.write(`${this.LOG_SOURCE} (_retrieveCurrentUserGroups) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Error);
            throw Error(`${err.message}`);
        }
    }

    private async _retrieveCurrentUser_AZGroupsMemeberOf(upn: string, siteUrl: string): Promise<void> {
        try{
            var endPoint = `https://cibctodayapps.phapps.cibc.com/api/ModernScript/${upn}/`;   // PROD
            
            // Check environment to set proper PH app REST API URL
            if (siteUrl.indexOf('cibcpte.sharepoint.com') > -1) {    
                endPoint = `https://uat-cibctodayapps.phapps.cibc.com/api/ModernScript/${upn}/`; // UAT    
            } else if (siteUrl.indexOf('/sit-') > -1) {              
                //TODO: Test purpose only. Have to remove upn for UAT, Prod
                // upn = "pte1.sharepoint@cibcpte.com";
                endPoint = `https://sit-cibctodayapps.phapps.cibc.com/api/ModernScript/${upn}/`; // sit
            } else if (siteUrl.indexOf('/dit-') > -1) {
                endPoint = `https://dit-cibctodayapps.phapps.cibc.com/api/ModernScript/${upn}/`; // dit
            } else if (siteUrl.indexOf('/dev8-') > -1) {                                
                endPoint = `https://dev8-cibctodayapps.dase-canc-toss-dev-shared.appserviceenvironment.net/api/ModernScript/${upn}/`; // dev8
            }
            Logger.writeJSON(`${this.LOG_SOURCE} - ${LogLevel.Info} ${JSON.stringify(endPoint)}`, LogLevel.Info);
            
            await this._httpClient
            .get(endPoint, HttpClient.configurations.v1)
            .then((res: HttpClientResponse): Promise<any> => {
                return res.json();
            })
            .then((response: any): void => {                
                Logger.write(`${this.LOG_SOURCE} _retrieveCurrentUser_AZGroupsMemeberOf - succeeded - ${LogLevel.Info}`, LogLevel.Verbose);
                Logger.writeJSON(`${JSON.stringify(response)}`, LogLevel.Verbose);

                // Use map to convert response into the internal array IAZGroupInfo
                this._user_AZGroupInfo = response.map((item:IAZGroupInfo) => {
                    return {              
                        displayName: item.displayName,
                        id: item.id,
                        description: item.description,
                        custom: item.custom,
                    }
                });
            });            
        } catch (err) {
            Logger.write(`${this.LOG_SOURCE} (_retrieveCurrentUser_AZGroupsMemeberOf) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Error);
            throw Error(`${err.message}`);
        }
    }
}