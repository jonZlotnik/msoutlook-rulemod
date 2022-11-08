import { NgModule } from "@angular/core";
import { BrowserModule } from "@angular/platform-browser";
import { MsalModule } from "@azure/msal-angular";
import { PublicClientApplication } from "@azure/msal-browser";
import AppComponent from "./app.component";

@NgModule({
  declarations: [AppComponent],
  imports: [
    BrowserModule,
    MsalModule.forRoot(
      new PublicClientApplication({
        auth: {
          clientId: "8ac9517d-d11f-4072-b1fd-c3d1bb234dc9", // Application (client) ID from the app registration
          authority: "https://login.microsoftonline.com/91c45569-c153-47f9-a0e2-abd39780f427", // The Azure cloud instance and the app's sign-in audience (tenant ID, common, organizations, or consumers)
          redirectUri: "http://localhost:4200", // This is your redirect URI
        },
        cache: {
          cacheLocation: "localStorage",
        },
      }),
      null,
      null
    ),
  ],
  bootstrap: [AppComponent],
})
export default class AppModule {}
