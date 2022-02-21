import * as React from 'react';
import styles from './MyAgenda.module.scss';
import * as strings from 'MyAgendaWebPartStrings';
import { IMyAgendaProps } from './IMyAgendaProps';
import { WebPartTitle } from '@pnp/spfx-controls-react/lib/WebPartTitle';
import { FontIcon, Link, Spinner, SpinnerSize, Text } from 'office-ui-fabric-react';

import { Agenda, Person, MgtTemplateProps } from '@microsoft/mgt-react/dist/es6/spfx';
import { Event as IEvent } from '@microsoft/microsoft-graph-types';
export default class MyAgenda extends React.Component<IMyAgendaProps, {}> {
  public render(): React.ReactElement<IMyAgendaProps> {
    return (
      <div className={styles.myAgenda}>
        <WebPartTitle displayMode={this.props.displayMode}
          title={this.props.description} className={styles.title}
          updateProperty={this.props.updateProperty}
          moreLink={() => {
            return <Link href='https://outlook.office.com/owa/?path=/calendar/view/Day' target='_blank'>{strings.ViewAll}</Link>;
          }}
        />
        <div className={styles.contentRow}>
          <Agenda days={1} showMax={6} date={new Date().toISOString()} >
            <Event template='event' data-type={"event"} />
            <NoData template='no-data' data-type={"no-data"} />
            <Loading template='loading' data-type={"loading"} />
          </Agenda>
        </div>
      </div>
    );
  }
}

export const Event = (props: MgtTemplateProps) => {
  const event: IEvent | undefined = props.dataContext ? props.dataContext.event : undefined;
  const startTime: Date = new Date(event.start.dateTime);
  const endTime: Date = new Date(event.end.dateTime);

  return (
    <div className={styles.mainWrapper}>
      <Link href={event.webLink} target='_blank'>
        <div className={`${styles.meetingWrapper} ${event.showAs}`}>
          <Text className={styles.infoContainer} block={true} nowrap={true}>
            <FontIcon iconName="Clock" />
            <span>
              {`${startTime.toLocaleTimeString()} - ${endTime.toLocaleTimeString()}`}
            </span>
            {
              event.location.displayName &&
              <span className={`${styles.spacer}`}>
                <FontIcon iconName="MapPin" />
                {event.location.displayName}
              </span>
            }
            <div className={styles.personas}>
              {
                event.attendees.map((attendee) => <Person className={styles.small} personQuery={attendee.emailAddress.address} />)
              }
            </div>
          </Text>
        </div>
        <div className={styles.subject}>{event.subject}</div>
      </Link>
    </div>
  );
};

export const NoData = (props: MgtTemplateProps) => {
  return (
    <div className={styles.noResults}>
      <img className={styles.noResultsImg} src={require<string>('../../assets/img_calendar_empty.svg')} alt="imgNoMeetings" />
      <div className={styles.noResultsText}>
        <span>{strings.NoMeetings}</span>
      </div>
    </div>
  );
};

export const Loading = (props: MgtTemplateProps) => {
  return (
    <div className={styles.noResults}>
      <Spinner label={strings.Loading} labelPosition="top" size={SpinnerSize.large} />
    </div>
  );
};