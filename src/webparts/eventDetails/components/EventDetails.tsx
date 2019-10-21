import * as React from 'react';
import styles from './EventDetails.module.scss';
import { IEventDetailsProps } from './IEventDetailsProps';
import * as strings from 'EventDetailsWebPartStrings';
import { sp, ItemAddResult } from "@pnp/sp";
import { PopupWindowPosition } from '@microsoft/sp-property-pane';

export default class EventDetails extends React.Component<IEventDetailsProps, {}> {
  
  private eventitems = false;
  private registerState = false;
  public render(): React.ReactElement<IEventDetailsProps> {
    let eventTitle;
    let eventLocation;
    let eventStart;
    let eventEnd;
    let self = this;

    if(this.props.registeredItem){
      self.registerState = true;
    }

    function ics(): void {
      var icsUrl = self.props.pageContext.web.absoluteUrl+'/_vti_bin/owssvr.dll?CS=109&Cmd=Display&List='+self.props.eventlistid+'&CacheControl=1&ID='+self.props.listItem+'&Using=event.ics';
      window.open(icsUrl);
    }

    function deleteItem():void {
      let list = sp.web.lists.getByTitle("EventRegistration");
      list.items.getById(self.props.registeredItem.Id).delete().then(_ => {
        self.registerState = false;
        location.reload();
      });
    }

    function singup():void {
      sp.web.lists.getByTitle("EventRegistration").items.add({
        Title: "Title",
        PersonId: self.props.userID,
        EventID: self.props.listItem
      }).then((iar: ItemAddResult) => {
          //console.info(iar); 
          self.registerState = true;
          location.reload();
      });
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
