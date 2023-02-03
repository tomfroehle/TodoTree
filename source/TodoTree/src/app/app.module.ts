import { BrowserModule } from '@angular/platform-browser';
import { BrowserAnimationsModule, NoopAnimationsModule } from '@angular/platform-browser/animations';
import { NgModule } from '@angular/core';

import { MatButtonModule } from '@angular/material/button';
import { MatToolbarModule } from '@angular/material/toolbar';
import { MatListModule } from '@angular/material/list';
import { MatMenuModule } from '@angular/material/menu';

import { AppRoutingModule } from './app-routing.module';
import { AppComponent } from './app.component';
import { HomeComponent } from './home/home.component';
import { ProfileComponent } from './profile/profile.component';

import { HTTP_INTERCEPTORS, HttpClientModule } from '@angular/common/http';
import { IPublicClientApplication, PublicClientApplication, InteractionType, BrowserCacheLocation, LogLevel } from '@azure/msal-browser';
import { MsalGuard, MsalInterceptor, MsalBroadcastService, MsalInterceptorConfiguration, MsalModule, MsalService, MSAL_GUARD_CONFIG, MSAL_INSTANCE, MSAL_INTERCEPTOR_CONFIG, MsalGuardConfiguration, MsalRedirectComponent } from '@azure/msal-angular';
import { FailedComponent } from './failed/failed.component';
import { TodoComponent } from './todo/todo.component';
import { Client } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProviderOptions, AuthCodeMSALBrowserAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser';

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

export function MSALInterceptorConfigFactory(): MsalInterceptorConfiguration {
  const protectedResourceMap = new Map<string, Array<string>>();
  protectedResourceMap.set('https://graph.microsoft.com/v1.0/me', ["user.read", "tasks.readWrite"]);

  return {
    interactionType: InteractionType.Redirect,
    protectedResourceMap
  };
}

export function MSALGuardConfigFactory(): MsalGuardConfiguration {
  return {
    interactionType: InteractionType.Redirect,
    authRequest: {
      scopes: ["user.read", "tasks.readWrite"]
    },
    loginFailedRoute: '/login-failed'
  };
}

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
    HomeComponent,
    ProfileComponent,
    FailedComponent,
    TodoComponent
  ],
  imports: [
    BrowserModule,
    BrowserAnimationsModule,
    AppRoutingModule,
    MatButtonModule,
    MatToolbarModule,
    MatListModule,
    MatMenuModule,
    HttpClientModule,
    MsalModule
  ],
  providers: [
    {
      provide: HTTP_INTERCEPTORS,
      useClass: MsalInterceptor,
      multi: true
    },
    {
      provide: MSAL_INSTANCE,
      useValue: publicClientApplication
    },
    {
      provide: PublicClientApplication,
      useValue: publicClientApplication
    },
    {
      provide: MSAL_GUARD_CONFIG,
      useFactory: MSALGuardConfigFactory
    },
    {
      provide: MSAL_INTERCEPTOR_CONFIG,
      useFactory: MSALInterceptorConfigFactory
    },
    {
      provide : Client,
      useFactory: GraphClient,
      deps: [MsalService, PublicClientApplication]
    },
    MsalService,
    MsalGuard,
    MsalBroadcastService
  ],
  bootstrap: [AppComponent, MsalRedirectComponent]
})
export class AppModule { }
