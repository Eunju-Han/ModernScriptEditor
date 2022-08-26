import { WebPartContext } from "@microsoft/sp-webpart-base";
import { spfi, SPFI, SPFx } from "@pnp/sp";
import { HttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { LogLevel, PnPLogging,
  Logger,
  ILogEntry,
  ILogListener,
} from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/batching";

var _sp: SPFI = null;
var _httpClient: HttpClient = null;

export const getSP = (context?: WebPartContext): SPFI => {
  if (_sp === null && context != null) {

    Logger.write(`(getSP) -${context.pageContext.web.absoluteUrl} ${LogLevel.Verbose}`, LogLevel.Verbose);
    _sp = spfi(context.pageContext.web.absoluteUrl).using(SPFx(context)).using(PnPLogging(LogLevel.Warning));       
  }
  return _sp;
};

export const getHttpClient = (context?: WebPartContext): HttpClient => {
  if (_httpClient === null && context != null) {

    Logger.write(`(getHttpClient) -${context.pageContext.web.absoluteUrl} ${LogLevel.Verbose}`, LogLevel.Verbose);
    _httpClient = context.httpClient;
  }
  return _httpClient;
};

export class CustomListener implements ILogListener {  
  log(entry: ILogEntry): void {

    if(entry.level >= Logger.activeLogLevel) {
      if (entry.level == LogLevel.Error && entry.level > Logger.activeLogLevel)
        console.log('%c' + entry.message, "color:red;");
      else if (entry.level == LogLevel.Warning)
        console.log('%c' + entry.message, "color:orange;");
      else if (entry.level == LogLevel.Info)
        console.log('%c'  + entry.message, "color:green;");
      else
        console.log('%c'+ entry.message, "color:blue;");
    }
  }
}
