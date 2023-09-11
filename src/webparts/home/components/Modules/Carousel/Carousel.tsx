import * as React from 'react';
import Container from 'react-bootstrap/Container';
import Row from 'react-bootstrap/Row';
import Col from 'react-bootstrap/Col';
import { ICarouselProps } from './ICarouselProps';
import { getListItems } from '../../../Services/SPOps';
import { config } from '../../../Services/Config';

export default class Carousel extends React.Component<ICarouselProps, any> {
    constructor(props: ICarouselProps) {
        super(props);
        this.state = {
            carouselData: []
        };
    }

    public async componentDidMount() {
        try {
            await getListItems(config.CarouselList, "*", "", "", "Order", this.props.context).then(async result => {
                this.setState({
                    carouselData: result
                });
            });
            console.log(this.state);
        } catch (error) {
            console.log(error);
        }
    }

    public render(): React.ReactElement<any> {
        return (
            <Row>
                Carousel
            </Row>
        )
    }
}