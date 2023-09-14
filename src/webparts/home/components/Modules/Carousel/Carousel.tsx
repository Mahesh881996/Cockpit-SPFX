import * as React from 'react';
import Container from 'react-bootstrap/Container';
import Row from 'react-bootstrap/Row';
import Col from 'react-bootstrap/Col';
import { ICarouselProps } from './ICarouselProps';
import { getListItems } from '../../../Services/SPOps';
import { config } from '../../../Services/Config';
import Carousel from 'react-bootstrap/Carousel';

export default class CarouselModule extends React.Component<ICarouselProps, any> {
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
                <Col md="12">
                    <Carousel>
                        {this.state.carouselData.map((data: any) => {
                            return <Carousel.Item>
                                <img className="d-block w-100" style={{height:"40vh"}} src={data.Image.Url} alt="First slide"/>
                                <Carousel.Caption>
                                    <h5>{data.Title}</h5>
                                    <p>{data.Description}</p>
                                </Carousel.Caption>
                            </Carousel.Item>
                        })}
                    </Carousel>
                </Col>
            </Row>
        )
    }
}