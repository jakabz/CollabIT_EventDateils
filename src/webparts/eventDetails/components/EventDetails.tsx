import * as React from 'react';
import styles from './EventDetails.module.scss';
import { IEventDetailsProps } from './IEventDetailsProps';
import * as strings from 'EventDetailsWebPartStrings';

export default class EventDetails extends React.Component<IEventDetailsProps, {}> {
  
  private eventitems = false;

  private singup(): void {
    alert('Hamarosan');
  }

  public render(): React.ReactElement<IEventDetailsProps> {
    let eventTitle;
    let eventLocation;
    let eventStart;
    let eventEnd;
    let self = this;

    function ics(): void {
      var icsUrl = self.props.pageContext.web.absoluteUrl+'/_vti_bin/owssvr.dll?CS=109&Cmd=Display&List='+self.props.eventlistid+'&CacheControl=1&ID='+self.props.listItem+'&Using=event.ics';
      window.open(icsUrl);
    }

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
            <button className={ styles.button }>
              <span className={ styles.label }>{strings.SingUp}</span>
            </button>
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
