import * as de from 'date-fns/locale/de';
import * as format from 'date-fns/format';
import * as React from 'react';
import styles from './EventList.module.scss';
import { Dialog, DialogFooter, DialogType } from 'office-ui-fabric-react/lib/Dialog';
import { find } from '@microsoft/sp-lodash-subset';
import { IEventListProps } from './IEventListProps';
import { Log } from '@microsoft/sp-core-library';
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';

export default class EventList extends React.Component<IEventListProps, { hideDialog: boolean; dialogTitle: string; dialogMessage: string; }> {

  /**
   *
   */
  constructor() {
    super();

    this.state = {
      hideDialog: true,
      dialogTitle: null,
      dialogMessage: null
    };
  }

  public render(): React.ReactElement<IEventListProps> {
    return (
      <div className={ styles.eventList }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <table>
              <tbody>
                <tr>
                  <th>Event</th>
                  <th>Start Date</th>
                  <th>End Date</th>
                  <th>Add to my calendar</th>
                </tr>
                { this.props.items &&
                  this.props.items.map(item => {
                    return <tr key={item.ID}>
                      <td>{item.Title}</td>
                      <td>{format(item.StartDate, 'dd, DD. MMMM YYYY HH:mm', { locale: de })}</td>
                      <td>{format(item.EndDate, 'dd, DD. MMMM YYYY HH:mm', { locale: de })}</td>
                      <td><a href={'/add/' + item.ID} onClick={ (e) => { e.preventDefault(); this.addToCalendar(item.ID); } }>join</a></td>
                    </tr>;
                  })
                }
              </tbody>
            </table>
            <Dialog hidden={this.state.hideDialog} onDismiss={this.closeDialog} dialogContentProps={{
              type: DialogType.largeHeader,
              title: this.state.dialogTitle,
              subText: this.state.dialogMessage
            }} modalProps={{isBlocking: false, containerClassName: 'ms-dialogMainOverride'}}>
              <DialogFooter>
                <PrimaryButton onClick={this.closeDialog} text="OK" />
              </DialogFooter>
            </Dialog>
          </div>
        </div>
      </div>
    );
  }

  public addToCalendar(itemId: number): void {
    const item = find(this.props.items, { ID: itemId });
    this.props.eventsService.addEventToCalendar(item)
      .then((calendarEvent) => {
        this.setState({
          hideDialog: false,
          dialogTitle: 'Neuer Termin erstellt',
          dialogMessage: `Termin '${calendarEvent.subject}' von ${format(calendarEvent.start.dateTime, 'dd, DD. MMMM YYYY HH:mm', { locale: de })} bis ${format(calendarEvent.end.dateTime, 'dd, DD. MMMM YYYY HH:mm', { locale: de })} wurde in deinem Kalender erstellt.`
        });
      })
      .catch(() => {
        this.setState({
          hideDialog: false,
          dialogTitle: 'Error',
          dialogMessage: 'Termin konnte nicht erstellt werden. Bitte wende dich an deinen Administrator.'
        });
      });
  }

  public closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }
}
