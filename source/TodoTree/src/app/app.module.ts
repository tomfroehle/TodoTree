import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule, NoopAnimationsModule } from '@angular/platform-browser/animations';
import { NgModule } from '@angular/core';

import { MatButtonModule } from '@angular/material/button';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatListModule } from '@angular/material/list';
import { MatMenuModule } from '@angular/material/menu';

import { AppComponent } from './app.component';
import { FailedComponent } from './failed/failed.component';
import { TodoComponent } from './todo/todo.component';
import { Client } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProviderOptions, AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';
import { BrowserCacheLocation, InteractionType, LogLevel, PublicClientApplication } from '@azure/msal-browser';
import { HttpClientModule } from '@angular/common/http';
import { MsalService, MsalModule, MSAL_INSTANCE, MsalBroadcastService, MsalRedirectComponent } from '@azure/msal-angular';

export function loggerCallback(logLevel: LogLevel, message: string) {
  console.log(message);
}

export const publicClientApplication = new PublicClientApplication({
    auth: {
      clientId: '8367ee60-5a62-4e0d-beb1-62d93087314e',
      authority: 'https://login.microsoftonline.com/common',
      redirectUri: '/',
      postLogoutRedirectUri: '/'
    },
    cache: {
      cacheLocation: BrowserCacheLocation.LocalStorage,
      storeAuthStateInCookie: false,
    },
    system: {
      loggerOptions: {
        loggerCallback,
        logLevel: LogLevel.Info,
        piiLoggingEnabled: false
      }
    }
  });

function GraphClient(authService: MsalService, clientApplication: PublicClientApplication): Client {

  const options: AuthCodeMSALBrowserAuthenticationProviderOptions = {
    account: authService.instance.getActiveAccount()!,
    interactionType: InteractionType.Popup,
    scopes: ["user.read", "tasks.readWrite"]
  };
  const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(clientApplication, options);


  // Initialize the Graph client
  const graphClient = Client.initWithMiddleware({
      authProvider
  });
  return graphClient;
}
@NgModule({
  declarations: [
    AppComponent,
    FailedComponent,
    TodoComponent
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    MatButtonModule,
    MatToolbarModule,
    MatListModule,
    MatMenuModule,
    HttpClientModule,
    MsalModule
  ],
  providers: [
    {
      provide: MSAL_INSTANCE,
      useValue: publicClientApplication
    },
    {
      provide: PublicClientApplication,
      useValue: publicClientApplication
    },
    {
      provide : Client,
      useFactory: GraphClient,
      deps: [MsalService, PublicClientApplication]
    },
    MsalService,
    MsalBroadcastService
  ],
  bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule { }
