import {
  IconButton,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack
} from "@fluentui/react";
import * as React from "react"

export interface AppProps {
  onSettingsClicked,
  onDownloadClicked,
  onSaveClicked: () => Promise<void>,
  onInsertClicked: () => Promise<void>,
  onInsertMarkClicked: () => Promise<void>,
  isOfficeWordEnvironment: boolean,
}

export interface AppState { 
  isSaving: boolean,
  isInserting: boolean,
}

export default class Footer extends React.Component<AppProps, AppState> {
  constructor(props, context) {
    super(props, context);
    this.state = {
      isSaving: false,
      isInserting: false,
    };
  }
  render() {
    return (
      <Stack className="footer" horizontal horizontalAlign="space-between">
        <Stack horizontal horizontalAlign="end" tokens={{ childrenGap: 5 }}>
          <IconButton title="Settings"
            iconProps={{ iconName: "Settings" }}
            onClick={this.props.onSettingsClicked} />
          <IconButton title="Download File"
            onClick={this.props.onDownloadClicked}
            iconProps={{ iconName: "DownloadDocument" }} />
          <IconButton title="Save"
            disabled={!this.props.isOfficeWordEnvironment}
            onClick={() => {
              this.setState({ isSaving: true });
              this.props.onSaveClicked().finally(() =>
                this.setState({ isSaving: false })
              );
            }}
            iconProps={{ iconName: "Save" }} />
          {this.state.isSaving ?
            <Spinner size={SpinnerSize.small} title="Saving" />
            : undefined}
        </Stack>
        <PrimaryButton text="Insert"
          disabled={this.state.isInserting ||
            !this.props.isOfficeWordEnvironment}
          split
          menuProps={{
            items: [{
              key: "insertMarked",
              text: "Insert Marked",
              onClick: () => {
                this.setState({ isInserting: true });
                this.props.onInsertMarkClicked().finally(() =>
                  this.setState({ isInserting: false })
                );
              },
            }],
            useTargetWidth: true,
            isBeakVisible: false,
          }}
          onClick={() => {
            this.setState({ isInserting: true });
            this.props.onInsertClicked().finally(() => 
              this.setState({ isInserting: false })
            );
          }}
          style={{ width: "68px" }} />
      </Stack>
    );
  }
}