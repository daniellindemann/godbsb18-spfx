import { IEventItem } from '../interfaces/IEventItem';
import { IEventsService } from '../interfaces/IEventsService';
import { IEventsServiceOptions } from '../interfaces/IEventsServiceOptions';
import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export class SPEventsService implements IEventsService {

  constructor(private context: WebPartContext, private options: IEventsServiceOptions) {
  }

  public get(): Promise<IEventItem[]> {
    const selects = [
      'ID',
      'Title',
      'StartDate',
      'EndDate'
    ];
    const filters = [
      `StartDate ge datetime'${new Date().toISOString()}'`
    ];
    const orders = [
      `StartDate asc`
    ];

    return new Promise<IEventItem[]>((resolve, reject) => {
      this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.options.listname}')/items?$select=${selects.join()}&$filter=${filters.join(' and ')}&$orderBy=${orders.join(',')}`, SPHttpClient.configurations.v1)
        .then((res: SPHttpClientResponse) => {
          return res.json();
        })
        .then((json) => {

          // error occured on query
          if(json.error) {
            reject(json.error);
            return;
          }

          resolve(json.value as IEventItem[]);
        });
    });
  }

  public addEventToCalendar(event: IEventItem): Promise<any> {
    if(!event.EndDate) {
      const d = new Date(event.StartDate as string);
      d.setHours(d.getHours() + 1);
      event.EndDate = d.toISOString();
    }

    return new Promise<any>((resolve, reject) => {
      this.context.msGraphClientFactory.getClient()
        .then((client) => {
          client.api('me/events')
            .version('v1.0')
            .post({
              subject: event.Title,
              start: {
                dateTime: event.StartDate instanceof Date ? (event.StartDate as Date).toISOString() : event.StartDate,
                timeZone: 'UTC'
              },
              end: {
                dateTime: event.EndDate instanceof Date ? (event.EndDate as Date).toISOString() : event.EndDate,
                timeZone: 'UTC'
              }
            })
            .then((res) => {
              resolve(res);
            })
            .catch((err) => {
              reject(err);
            });
        });
      });
  }
}
