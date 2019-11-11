import * as React from 'react';
import styles from './EventDetails.module.scss';
import { IEventDetailsProps } from './IEventDetailsProps';
import * as strings from 'EventDetailsWebPartStrings';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

export default class EventDetails extends React.Component<IEventDetailsProps, {}> {
  
  private eventitems = false;
  private registerState = false;

  public render(): React.ReactElement<IEventDetailsProps> {
    let eventTitle;
    let eventLocation;
    let eventStart;
    let eventEnd;
    let self = this;

    let newReg = function(title,userid,eventid): Promise<HttpClientResponse> {
      const postURL = 'https://prod-26.westeurope.logic.azure.com:443/workflows/5a1b49e8e3414ee09ef8ff73a8f1935a/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=5rDqt_1m4rssK2XF2cWZbnadgvmIAXcl5rc-NCILUjU';
      const body: string = JSON.stringify({
        "eventTitle": title,
        "userEmail": self.props.pageContext.user.loginName, //"eszter.koscsak.admin@QualysoftHolding.onmicrosoft.com",
        "eventID": eventid
      });
      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Content-type', 'application/json');
      const httpClientOptions: IHttpClientOptions = {
        body: body,
        headers: requestHeaders
      };
  
      console.log("List item creating.");
  
      return self.props.context.httpClient.post(
        postURL,
        HttpClient.configurations.v1,
        httpClientOptions)
        .then((response: Response): Promise<HttpClientResponse> => {
          console.log("List item created.");
          console.info(response.json());
          location.reload();
          return response.json();
        });
    };
  
    let delReg = function(regid): Promise<HttpClientResponse> {
      const postURL = 'https://prod-15.westeurope.logic.azure.com:443/workflows/f6e8465e65b44943a72294f045b554ef/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=uoTsVGqwU5_tpSnOTIxXUQxSoEAndvE2f9MmlgrjC_4';
      const body: string = JSON.stringify({
        "eventID": regid
      });
      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Content-type', 'application/json');
      const httpClientOptions: IHttpClientOptions = {
        body: body,
        headers: requestHeaders
      };
  
      console.log("List item deleting.");
  
      return self.props.context.httpClient.post(
        postURL,
        HttpClient.configurations.v1,
        httpClientOptions)
        .then((response: Response): Promise<HttpClientResponse> => {
          console.log("List item deleted.");
          console.info(response.json());
          location.reload();
          return response.json();
        });
    };

    if(this.props.registeredItem){
      self.registerState = true;
    }

    function ics(): void {
      var icsUrl = self.props.pageContext.web.absoluteUrl+'/_vti_bin/owssvr.dll?CS=109&Cmd=Display&List='+self.props.eventlistid+'&CacheControl=1&ID='+self.props.listItem+'&Using=event.ics';
      window.open(icsUrl);
    }

    function deleteItem():void {
      var msg = delReg(self.props.registeredItem.Id);
      console.info(msg);
    }

    function singup():void {
      var msg = newReg(eventTitle,self.props.userID,self.props.listItem);
      console.info(msg);
    }

    //console.info(this.props.registeredItem);

    if(this.props.listItem){
      this.props.listsItems.forEach(item => {
        if(item.Id === this.props.listItem){
          var options = { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' };
          this.eventitems = true;
          eventTitle = item.Title;
          eventLocation = item.Location;
          if(item.EventDate){
            var d = new Date(item.EventDate);
            eventStart = d.toLocaleString('en-US',options);
          } else {
            eventStart = "";
          }
          if(item.EndDate){
            var n = new Date(item.EndDate);
            eventEnd = n.toLocaleString('en-US',options);
          } else {
            eventEnd = "";
          }
        }
      });
    } else {
      this.eventitems = false;
    }
    
    return (
      <div>
      { this.eventitems ? 
        <div className={ styles.eventDetails }>
          { this.props.TitleViewBool ? <h3 className={ styles.title }>{ eventTitle }</h3> : ""}
          { this.props.LocationViewBool ? <div className={ styles.location }> ({ eventLocation }) </div> : ""}
          <div>
            { this.props.StartViewBool ? <span>{ eventStart } </span> : <span></span>}
            { this.props.EndViewBool ? <span>- { eventEnd }</span> : <span></span>}
          </div>
          { this.props.RegisterButtonViewBool || this.props.SaveEventButtonViewBool ?
          <div className={styles.buttonBox}>
            { this.props.RegisterButtonViewBool ?
            <span>
              { self.registerState ?
                <button className={ styles.button }  onClick={deleteItem}>
                  <span className={ styles.label }>{strings.UnSubscribe}</span>
                </button>
              : <button className={ styles.button }  onClick={singup}>
                  <span className={ styles.label }>{strings.SingUp}</span>
                </button>
              }
            </span>
            : "" }
            { this.props.SaveEventButtonViewBool ?
            <button className={ styles.button } onClick={ics}>
              <span className={ styles.label }>{ strings.SaveEvent }</span>
            </button>
            : "" }
          </div> : ""}
        </div>
      : <div className={ styles.setWP }>Please edit this web part!</div>}
      </div>
    );
  }
}
