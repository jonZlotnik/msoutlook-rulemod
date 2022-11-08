import { Component } from "@angular/core";
import { MsalService } from "@azure/msal-angular";
import { InteractionType, PublicClientApplication } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import { AuthCodeMSALBrowserAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";

@Component({
  selector: "app-home",
  templateUrl: "./app.component.html",
})
export default class AppComponent {
  public authenticated: boolean;
  public graphClient?: Client;
  private accessToken = null;
  welcomeMessage = "Welcome";

  title = "Outlook Rule Tools";
  isIframe = false;
  loginDisplay = false;

  constructor(private authService: MsalService) {}

  login() {
    this.authService.loginPopup().subscribe({
      next: (result) => {
        // eslint-disable-next-line no-undef
        console.log(result);
        this.setLoginDisplay();
      },
      // eslint-disable-next-line no-undef
      error: (error) => console.log(error),
    });
  }

  setLoginDisplay() {
    this.loginDisplay = this.authService.instance.getAllAccounts().length > 0;
  }

  async run() {
    // eslint-disable-next-line no-undef
    this.isIframe = window !== window.parent && !window.opener;
     const authProvider = new AuthCodeMSALBrowserAuthenticationProvider(
      this.authService.instance as PublicClientApplication,
      {
        account: this.authService.instance.getActiveAccount()!,
        scopes: ["MailboxSettings.Read"],
        interactionType: InteractionType.Popup,
      }
    );
    this.graphClient = Client.initWithMiddleware({ authProvider: new OfficeAuthProvider() });
  }
}
