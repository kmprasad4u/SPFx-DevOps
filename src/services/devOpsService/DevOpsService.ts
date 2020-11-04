import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { AadHttpClientFactory, AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import { SPHttpClient } from "@microsoft/sp-http";

import { IDevOpsService } from './IDevOpsService';

export class DevOpsService implements IDevOpsService {

  public static readonly serviceKey: ServiceKey<IDevOpsService> = ServiceKey.create<IDevOpsService>('SPFx:DevOpsService', DevOpsService);

  private _aadHttpClientFactory : AadHttpClientFactory;
  private _spHttpClient : SPHttpClient;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._aadHttpClientFactory =  serviceScope.consume(AadHttpClientFactory.serviceKey);
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
    });
  }

  public getProjects1() : void { 
    // Token gets retrieved and 
    // while trying to call the API, token is being passed to API Url
    // But redirect is happening to login page and during that time, getting below message

    // Uncaught (in promise) SyntaxError: Unexpected token < in JSON at position 4

    
      this._aadHttpClientFactory.getClient("https://app.vssps.visualstudio.com").then((client: AadHttpClient) => {      
        client.get(`https://app.vssps.visualstudio.com/_apis/accounts`, AadHttpClient.configurations.v1)
        .then((response: HttpClientResponse) => {
          console.log(["Try1", response]);
          return response.json();
        })
        .then((projects: any): void => {
          console.log(["Try1", projects]);
        });
      });
  }

  public getProjects2() : void {     

    // Token gets retrieved and 
    // while trying to call the API, token is being passed to API Url
    // But redirect is happening to login page and during that time, getting below message

    // Uncaught (in promise) SyntaxError: Unexpected token < in JSON at position 4

    this._aadHttpClientFactory.getClient("https://app.vssps.visualstudio.com").then((client: AadHttpClient) => {      
      client.get(`https://dev.azure.com/spfxlab/_apis/projects?api-version=6.0`, AadHttpClient.configurations.v1)
      .then((response: HttpClientResponse) => {
        console.log(["Try2", response]);
        return response.json();
      })
      .then((projects: any): void => {
        console.log(["Try2", projects]);
      });
    });    
  }

  public getProjects3() : void {  
    // Token is not getting retrieved

    // Getting below error messages

    // Cannot read property 'getRequestInfo' of null 
    //    parseTokenFromUrl @ spfxsinglesignon.aspx:38
    //    onload @ VM1521 spfxsinglesignon.aspx:58

    // The resource principal named https://dev.azure.com was not found in the tenant named 4283xxxx-xxxx-xxxx-xxxx-xxxx0562e6b. 
    // This can happen if the application has not been installed by the administrator of the tenant or consented to by any user 
    // in the tenant. You might have sent your authentication request to the wrong tenant.
    this._aadHttpClientFactory.getClient("https://dev.azure.com").then((client: AadHttpClient) => {      
      client.get(`https://dev.azure.com/spfxlab/_apis/projects?api-version=6.0`, AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        console.log(["Try3", response]);
        return response.json();
      })
      .then((projects: any): void => {
        console.log(["Try3", projects]);
      });
    });
  }
  
  
  public getProjects4() : void {  

    
    // Uncaught (in promise) Error: interaction_required: AADSTS65001: The user or administrator has not consented to use the application 
    // with ID 'c91exxxx-xxxx-xxxx-xxxx-xxxxfbce5ba7' named 'sp-fx-dev-ops-client-side-solution'. Send an interactive authorization 
    // request for this user and resource.
    // at chunk.spoImplicit_none_592577b526a7e5da04e4.js:1

    this._aadHttpClientFactory.getClient("c91exxxx-xxxx-xxxx-xxxx-xxxxfbce5ba7").then((client: AadHttpClient) => {      
      client.get(`https://app.vssps.visualstudio.com/_apis/accounts`, AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        console.log(["Try4", response]);
        return response.json();
      })
      .then((projects: any): void => {
        console.log(["Try4", projects]);
      });
    });
  }

  public getProjects5() : void {  

    // Uncaught (in promise) Error: interaction_required: AADSTS65001: The user or administrator has not consented to use the application 
    // with ID 'c91exxxx-xxxx-xxxx-xxxx-xxxxfbce5ba7' named 'sp-fx-dev-ops-client-side-solution'. Send an interactive authorization 
    // request for this user and resource.
    // at chunk.spoImplicit_none_592577b526a7e5da04e4.js:1

    this._aadHttpClientFactory.getClient("c91exxxx-xxxx-xxxx-xxxx-xxxxfbce5ba7").then((client: AadHttpClient) => {      
      client.get(`https://dev.azure.com/spfxlab/_apis/projects?api-version=6.0`, AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        console.log(["Try5", response]);
        return response.json();
      })
      .then((projects: any): void => {
        console.log(["Try5", projects]);
      });
    });
  }

  public getProjects6() : void {  

    // 499b84ac-1321-427f-aa17-267ca6975798 - is Azue DevOps
    this._aadHttpClientFactory.getClient("499b84ac-1321-427f-aa17-267ca6975798").then((client: AadHttpClient) => {      
      client.get(`https://dev.azure.com/spfxlab/_apis/projects?api-version=6.0`, AadHttpClient.configurations.v1).then((response: HttpClientResponse) => {
        console.log(["Try5", response]);
        return response.json();
      })
      .then((projects: any): void => {
        console.log(["Try5", projects]);
      });
    });
  }
}
  