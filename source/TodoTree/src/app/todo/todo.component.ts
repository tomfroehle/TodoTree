import { Component, OnInit } from '@angular/core';
import { MsalService } from '@azure/msal-angular';
import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { TodoTaskList } from "@microsoft/microsoft-graph-types";

import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import { Client } from "@microsoft/microsoft-graph-client";
import { Observable } from 'rxjs';

@Component({
  selector: 'app-todo',
  templateUrl: './todo.component.html',
  styleUrls: ['./todo.component.css']
})
export class TodoComponent implements OnInit {

  output = {};
  taskLists: TodoTaskList[] = [];

  constructor(private authService: MsalService, private clientApplication: PublicClientApplication) { }

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

    var result = await graphClient.api('/me/todo/lists?$expand=tasks').get();

    var taskLists = result.value as TodoTaskList[];
    for (const taskList of taskLists) {
      var result2 = await graphClient.api(`/me/todo/lists/${taskList.id}/tasks`).get();
      taskList.tasks = result2.value;
    }

    this.taskLists = taskLists;
}
}
