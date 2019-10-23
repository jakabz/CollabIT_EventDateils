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
      const postURL = 'https://prod-117.westeurope.logic.azure.com:443/workflows/9bc942ec999541e1a9d3293d4d5a20b6/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=RnCYjU9qinlbOk1OhjI3faGu4cxxXEHDwpzH3ZiEAkI';
      const body: string = JSON.stringify({
        "eventTitle": title,
        "userEmail": "eszter.koscsak.admin@QualysoftHolding.onmicrosoft.com",
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
      const postURL = 'https://prod-58.westeurope.logic.azure.com:443/workflows/2c80a7736fe64a60b1d2aaca1df2ba42/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=uopoe6zNaUjDA1TWx5MZskqe3XZi6Y1LdBJkqOw8p94';
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
      : <div className={ styles.setWP }>Please setting webpart!</div>}
      </div>
    );
  }
}
