import { NgModule } from "@angular/core";
import { RouterModule, Routes } from "@angular/router";
import { BrowserUtils } from "@azure/msal-browser";
import AppComponent from "./app.component";

const routes: Routes = [
  {
    path: "",
    component: AppComponent,
  },
];

// eslint-disable-next-line no-undef
const isIframe = window !== window.parent && !window.opener;

@NgModule({
  imports: [
    RouterModule.forRoot(routes, {
      // Don't perform initial navigation in iframes or popups
      initialNavigation: !BrowserUtils.isInIframe() && !BrowserUtils.isInPopup() ? "enabledNonBlocking" : "disabled", // Set to enabledBlocking to use Angular Universal
    }),
  ],
  exports: [RouterModule],
})
export class AppRoutingModule {}
