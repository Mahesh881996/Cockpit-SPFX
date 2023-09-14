import * as React from 'react';
import styles from './Home.module.scss';
import { IHomeProps } from './IHomeProps';
import { escape } from '@microsoft/sp-lodash-subset';
import Container from 'react-bootstrap/Container';
import Row from 'react-bootstrap/Row';
import Col from 'react-bootstrap/Col';
import CarouselModule from './Modules/Carousel/Carousel';

export default class Home extends React.Component<IHomeProps, {}> {
  public render(): React.ReactElement<IHomeProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section>
        <Container fluid>
          <CarouselModule context={this.props.context}/>
        </Container>
      </section>
    );
  }
}
