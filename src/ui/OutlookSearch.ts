import Component = Coveo.Component;
import ComponentOptions = Coveo.ComponentOptions;
import IComponentBindings = Coveo.IComponentBindings;
import Facet = Coveo.Facet;
import $$ = Coveo.$$;
import QueryEvents = Coveo.QueryEvents;
import IBuildingQueryEventArgs = Coveo.IBuildingQueryEventArgs;
import Initialization = Coveo.Initialization;

declare var MicrosoftOAuth2Provider: any;
declare var OAuth2Client: any;

export class OutlookSearch extends Component {
  static ID = 'OutlookSearch';

  constructor(public element: HTMLElement, public options: any, public bindings: IComponentBindings) {
    super(element, OutlookSearch.ID, bindings);
    this.options = ComponentOptions.initComponentOptions(element, OutlookSearch, options);

    $$(this.element).text(this.options.dummyOptionText);
    this.bind.onRootElement(QueryEvents.doneBuildingQuery, (args: IBuildingQueryEventArgs)=> this.doneBuildingQuery(args));
  }

  public enable() {
    super.enable();

    setTimeout(()=>{
      this.getOpenIdConfig();
    }, 1);
  }

  public disable() {
    super.disable();
  }

  private getOpenIdConfig() {

    let azAccessToken = document.cookie.replace(/(?:(?:^|.*;\s*)azAccessToken\s*\=\s*([^;]*).*$)|^.*$/, "$1");

    if (MicrosoftOAuth2Provider.hasAccessToken()) {
      let uri = MicrosoftOAuth2Provider.refreshToken();

      let a = MicrosoftOAuth2Provider.getAccessToken();
      setTimeout(()=>{
        this.getEmails();
      }, 1);
      return;
    }
    this.requestOAuthToken();
  }

  private requestOAuthToken() {
    // Create a new request
    let request = new OAuth2Client.Request({
        client_id: '<need client id>',  // required
        redirect_uri: 'https://localhost:8080/index.html',
        response_mode: 'fragment',
        scope: 'openid profile https://outlook.office.com/mail.readwrite',
        state: 12345
    });

    // Give it to the provider
    let uri = MicrosoftOAuth2Provider.requestToken(request);

    // Later we need to check if the response was expected so save the request
    MicrosoftOAuth2Provider.remember(request);

    // Do the redirect
    window.location.href = uri;
  }

  private renderEmail(v) {
    return `<div class="emailBox">
      <div class="date">${new Date(v.ReceivedDateTime).toLocaleString()}</div>
      <div class="from">From: ${v.From.EmailAddress.Name} </div>
      <div class="subject"><a class="CoveoResultLink" href="${v.WebLink}" target="outlook">${v.Subject}</a></div>
      ${v.BodyPreview}
    </div>`;
  }

  private getEmails() {

    let token = MicrosoftOAuth2Provider.getAccessToken(), q = this.queryStateModel.attributes['q'];
    if (q) {
      q = '$search=' + encodeURIComponent('"' + q + '"');
    }
    else {
      q = '$top=10';
    }

    var xmlHttpRequest = new XMLHttpRequest();
    xmlHttpRequest.open('GET', 'https://outlook.office.com/api/v2.0/me/messages?' + q, true);
    xmlHttpRequest.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
    xmlHttpRequest.setRequestHeader('Authorization', 'Bearer ' + token);
    xmlHttpRequest.setRequestHeader('X-AnchorMailbox', 'jdevost@coveo.com');
    xmlHttpRequest.onload = ()=>{
      let json = JSON.parse(xmlHttpRequest.responseText), html = [];
      html = json.value.map( this.renderEmail );
      $$( $$(this.element).find('.CoveoResultList') ).setHtml(html.join(''));
    };

    xmlHttpRequest.onerror = (e)=>{
      console.log('There was an error!',e);
      this.requestOAuthToken();
    };

    xmlHttpRequest.send();
  }

  private doneBuildingQuery(args: IBuildingQueryEventArgs) {
    this.getEmails();
  }
}

Initialization.registerAutoCreateComponent(OutlookSearch);
