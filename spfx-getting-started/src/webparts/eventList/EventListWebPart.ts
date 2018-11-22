import * as React from 'react';
import * as ReactDom from 'react-dom';
import * as strings from 'EventListWebPartStrings';
import EventList from './components/EventList';
import { BaseClientSideWebPart, IPropertyPaneConfiguration, PropertyPaneTextField } from '@microsoft/sp-webpart-base';
import {
  Environment,
  EnvironmentType,
  Log,
  Version
  } from '@microsoft/sp-core-library';
import { IEventItem } from '../../interfaces/IEventItem';
import { IEventListProps } from './components/IEventListProps';
import { IEventsService } from '../../interfaces/IEventsService';
import { MockEventsService } from '../../services/MockEventsService';
import { MSGraphClient } from '@microsoft/sp-http';
import { SPEventsService } from '../../services/SPEventsService';


export interface IEventListWebPartProps {
  listname: string;
}

export default class EventListWebPart extends BaseClientSideWebPart<IEventListWebPartProps> {

  private static EventListWebPartSource: string = 'EventListWebPart';

  private eventsService: IEventsService = null;

  protected onInit(): Promise<void> {
    // get's only called once after page load
    this.eventsService = Environment.type == EnvironmentType.Local ?
      new MockEventsService() :
      new SPEventsService(this.context, this.properties);
    return Promise.resolve();
  }

  public render(): void {
    if(!this.properties.listname) {
      this.context.statusRenderer.renderError(this.domElement, 'Configure the list that contains the event data via webpart properties.');
      return;
    }

    this.eventsService.get().then((events) => {
      Log.info(EventListWebPart.EventListWebPartSource, `Got ${events ? events.length : 0} events`, this.context.serviceScope);

      const element: React.ReactElement<IEventListProps> = React.createElement(
        EventList,
        {
          eventsService: this.eventsService,
          items: events
        }
      );
      ReactDom.render(element, this.domElement);
    })
    .catch((err) => {
      this.context.statusRenderer.renderError(this.domElement, err.message ? err.message : 'Unable to get event data');
      Log.warn(EventListWebPart.EventListWebPartSource, err.message, this.context.serviceScope);
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('listname', {
                  label: strings.ListNameLabel,
                  description: strings.ListNameDescription,
                  placeholder: strings.ListNamePlaceholder
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
