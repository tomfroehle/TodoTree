import { Component, OnInit, Inject, OnDestroy } from '@angular/core';
import { MsalService, MsalBroadcastService, MSAL_GUARD_CONFIG, MsalGuardConfiguration } from '@azure/msal-angular';
import { AuthenticationResult, InteractionStatus, PopupRequest, RedirectRequest, EventMessage, EventType } from '@azure/msal-browser';
import { Subject } from 'rxjs';
import { filter, takeUntil } from 'rxjs/operators';

@Component({
  selector: 'app-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent implements OnInit {
  constructor(private authService: MsalService, private msalBroadcastService: MsalBroadcastService) {}
  isLoggedIn = false;

  ngOnInit(): void {
    this.msalBroadcastService.inProgress$
    .pipe(
      filter((status: InteractionStatus) => status === InteractionStatus.None),
    )
    .subscribe(() => {
      this.isLoggedIn = this.checkAndSetActiveAccount();
    })
  }

  checkAndSetActiveAccount() : boolean {
    const activeAccount = this.authService.instance.getActiveAccount();
    const allAccounts = this.authService.instance.getAllAccounts();
    if (!activeAccount && allAccounts.length > 0) {
      this.authService.instance.setActiveAccount(allAccounts[0]);
    }

    return this.authService.instance.getActiveAccount() !== null;
  }

  login() {
    this.authService.instance.loginRedirect();
  }
}
