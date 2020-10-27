import * as React from "react";
import { Provider, Flex, Text, Button, Header } from "@fluentui/react-northstar";

import { teamsTheme, teamsDarkTheme, teamsHighContrastTheme } from "@fluentui/react-northstar";
import { ThemePrepared } from "@fluentui/styles";

import TeamsBaseComponent, { ITeamsBaseComponentState } from "../../TeamsBaseComponent";
import * as microsoftTeams from "@microsoft/teams-js";
/**
 * State for the testTabTab React component
 */
export interface ITestTabState extends ITeamsBaseComponentState {
    entityId?: string;
}

/**
 * Properties for the testTabTab React component
 */
export interface ITestTabProps {

}

/**
 * Implementation of the test Tab content page
 */
export class TestTab extends TeamsBaseComponent<ITestTabProps, ITestTabState> {

    public async componentWillMount() {
        this.updateTheme(this.getQueryVariable("theme"));


        if (await this.inTeams()) {
            microsoftTeams.initialize();
            microsoftTeams.registerOnThemeChangeHandler(this.updateTheme);
            microsoftTeams.getContext((context) => {
                microsoftTeams.appInitialization.notifySuccess();
                this.setState({
                    entityId: context.entityId
                });
                this.updateTheme(context.theme);
            });
        } else {
            this.setState({
                entityId: "This is not hosted in Microsoft Teams"
            });
        }
    }

    /**
     * The render() method to create the UI of the tab
     */
    public render() {
        return (
            <Provider theme={this.state.theme}>
                <Flex fill={true} column styles={{
                    padding: ".8rem 0 .8rem .5rem"
                }}>
                    <Flex.Item>
                        <Header content="This is your tab" />
                    </Flex.Item>
                    <Flex.Item>
                        <div>

                            <div>
                                <Text content={this.state.entityId} />
                            </div>

                            <div>
                                <Button onClick={() => alert("It worked!")}>A sample button</Button>
                            </div>
                        </div>
                    </Flex.Item>
                    <Flex.Item styles={{
                        padding: ".8rem 0 .8rem .5rem"
                    }}>
                        <Text size="smaller" content="(C) Copyright home" />
                    </Flex.Item>
                </Flex>
            </Provider>
        );
    }

    protected updateTheme = (themeStr?: string): void => {
        let theme: ThemePrepared<any>;
        switch (themeStr) {
            case "dark":
                theme = teamsDarkTheme;
                break;
            case "contrast":
                theme = teamsHighContrastTheme;
                break;
            case "default":
            default:
                theme = teamsTheme;
        }
        this.setState({ theme });
    }
}
