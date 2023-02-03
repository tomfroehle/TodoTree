import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { InteractionType, IPublicClientApplication, PublicClientApplication } from "@azure/msal-browser";

import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Client } from "@microsoft/microsoft-graph-client";
import { GRAPH_ENDPOINT } from '../profile/profile.component';
import { HttpClient } from '@angular/common/http';
import { firstValueFrom } from 'rxjs';

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {

  output = {};
  output2 = {};

  constructor(private http: HttpClient, private authService: MsalService, private clientApplication: PublicClientApplication) { }

  async ngOnInit(): Promise<void> {

    const options: AuthCodeMSALBrowserAuthenticationProviderOptions = {
        account: this.authService.instance.getActiveAccount()!, // the AccountInfo instance to acquire the token for.
        interactionType: InteractionType.Popup, // msal-browser InteractionType
        scopes: ["user.read", "tasks.readWrite"] // example of the scopes to be passed
    };

    // Pass the PublicClientApplication instance from step 2 to create AuthCodeMSALBrowserAuthenticationProvider instance
    const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(this.clientApplication, options);


    // Initialize the Graph client
    const graphClient = Client.initWithMiddleware({
        authProvider
    });
    try {
      let userDetails = await graphClient.api('/me').get();
      this.output = userDetails;
    } catch (error) {
      throw error;
    }

    this.output2 = await firstValueFrom(this.http.get('https://graph.microsoft.com/v1.0/me/todo/lists'));
  }

}
