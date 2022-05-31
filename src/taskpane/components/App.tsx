import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Header from "./Header";
import HeroList, { HeroListItem } from "./HeroList";
import Progress from "./Progress";
import { getDateFromTimeSelection } from "@fluentui/date-time-utilities";

/* global require */
let mailItemMarkingProperties: any;

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  listItems: HeroListItem[];
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      listItems: [],
    };

    Office.context.mailbox.item.loadCustomPropertiesAsync(
      (asyncResult) => (mailItemMarkingProperties = asyncResult.value)
    );
  }

  componentDidMount() {
  }

  saveProperties = async () => {
    this.setMailItemMarkingProperties();
  };

  loadProperties = async () => {
    var properties = mailItemMarkingProperties.getAll();
    this.setState({
      listItems: [
        {
          icon: "Ribbon",
          primaryText: "Property loaded: " + properties["test"],
        }
      ],
    });
    console.log(properties);
  };

  setMailItemMarkingProperties() {
    this.clearMailItemMarkingProperties((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("Error clearing properties");
      } else {
        var currentDateTime = new Date();
        this.setState({
          listItems: [
            {
              icon: "Ribbon",
              primaryText: "Property saved at: " + currentDateTime,
            }
          ],
        });
        mailItemMarkingProperties.set("test", currentDateTime);
        mailItemMarkingProperties.saveAsync((asyncResult2) => {
          if (asyncResult2.status === Office.AsyncResultStatus.Failed) {
            console.error("Error saving properties");
          }
        });
      }
    });
  }

  clearMailItemMarkingProperties(callback: any) {
    var properties = mailItemMarkingProperties.getAll();
    if (properties["name"]) {
      properties["name"].forEach((item) => {
        mailItemMarkingProperties.remove(item);
      });
    }
    mailItemMarkingProperties.saveAsync(callback);
  }

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <HeroList message="Sample add-in to test saving custom properties" items={this.state.listItems}>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.saveProperties}>
            Save Properties
          </DefaultButton>
          <br></br>
          <DefaultButton className="ms-welcome__action" iconProps={{ iconName: "ChevronRight" }} onClick={this.loadProperties}>
            Load Properties
          </DefaultButton>
        </HeroList>
      </div>
    );
  }
}
