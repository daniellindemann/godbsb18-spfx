import { IEventItem } from '../../../interfaces/IEventItem';
import { IEventsService } from '../../../interfaces/IEventsService';

export interface IEventListProps {
  eventsService: IEventsService;
  items: IEventItem[];
}
