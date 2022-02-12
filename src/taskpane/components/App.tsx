import * as React from "react"
import {
  IDropdownOption,
  Panel, PanelType,
  Stack,
  Spinner,
  SpinnerSize,
  MessageBar,
  MessageBarType,
} from "@fluentui/react";
import Settings from "./Settings";
import { OptionsSaver } from "../saver";
import { LineNumberType } from "./SettingsOptionsLists";
import Header from "./Header";
import Footer from "./Footer";
import SettingsPanelHeader from "./SettingsPanelHeader";

const environmentWarningMessage =
  "This page is not opened in a Micorsoft Office Word document." +
  " Some features of add-in may not be availabel";

const messageBarCloseDelay = 1000;

export interface AppProps {
  isOfficeInitialized: boolean,
  isOfficeWordEnvironment: boolean,
  editorContainerID: string,
  supportLanguages: IDropdownOption[],
  modelLanguageId: string,
  buttonMarkHandler: () => void,
  buttonUnmarkHandler: () => void,
  buttonDownloadHandler: () => void,
  buttonSaveHandler: () => Promise<void>,
  buttonInsertHandler: () => Promise<void>,
  buttonInsertMarkedHandler: (boolean?) => Promise<void>,
  buttonOpenHandler: () => void,
  languageChangeHandler: (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number) => void,
  encodingChangeHandler: (string, boolean) => void,
}

export interface AppState {
  isSettingsOpen: boolean,
  settingPanelTabId: string,
  environmentWarningMessageBar: boolean,
  showMessage: boolean,
  messageType: MessageBarType,
  message: string,
  commonLanguages: string[],
}

export default class App extends React.Component<AppProps, AppState> {
  settingsPanelRef = React.createRef<Settings>();
  constructor(props, context) {
    super(props, context);
    this.state = {
      isSettingsOpen: false,
      settingPanelTabId: "tab1",
      environmentWarningMessageBar: true,
      showMessage: false,
      messageType: MessageBarType.info,
      message: "",
      commonLanguages: [],
    };
  }

  componentDidMount(): void {
    new OptionsSaver().getDocOptions(
      this.props.isOfficeWordEnvironment).then(({ commonLanguages }) => {
        this.setState({ commonLanguages });
      });
  }

  onChangeLanguage = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number) => {
    if (option.key === "manage") {
      this.openSettingsPanel("tab3");
    } else {
      this.props.languageChangeHandler(event, option, index);
    }
  }

  onInsertCode = async () => {
    return this.props.buttonInsertHandler().then(() => {
      this.showMessage(MessageBarType.success,
        "Insert code successfully!",
        messageBarCloseDelay);
    }).catch(() => {
      this.showMessage(MessageBarType.error,
        "Insert code failed!",
        messageBarCloseDelay);
    });
  }

  onInsertMarkedCode = async () => {
    const { lineNumberType } = await (new OptionsSaver())
      .getDocOptions(this.props.isOfficeWordEnvironment);
    return this.props.buttonInsertMarkedHandler(
      lineNumberType === LineNumberType.Reorder
    ).then(() => {
      this.showMessage(MessageBarType.success,
        "Insert code successfully!",
        messageBarCloseDelay);
    }).catch(() => {
      this.showMessage(MessageBarType.error,
        "Insert code failed!",
        messageBarCloseDelay);
    });
  }

  onSaveCode = () => {
    return this.props.buttonSaveHandler().then(() => {
      this.showMessage(MessageBarType.success,
        "Save code successfully!",
        messageBarCloseDelay);
    }).catch(() => {
      this.showMessage(MessageBarType.error,
        "Save code failed!",
        messageBarCloseDelay);
    });
  }

  onRenderSettingsPanelHeader = () => (
    <SettingsPanelHeader settingsPanelRef={this.settingsPanelRef}
      onEncodingChanged={this.props.encodingChangeHandler}
      onCloseSettingsPanel={this.closeSettingsPanel}
      showMessage={this.showMessage}
      isOfficeWordEnvironment={this.props.isOfficeWordEnvironment}/>
  )
  
  openSettingsPanel = (tabId?: string) => {
    let state = { isSettingsOpen: true };
    if (tabId) state["settingPanelTabId"] = tabId;
    this.setState(state);
  }

  closeSettingsPanel = () => {
    this.setState({
      settingPanelTabId: "tab1",
      isSettingsOpen: false,
    });
  }

  showMessage = (
    messageType: MessageBarType,
    message: string,
    closeDelay?: number) => {
    this.setState({
      showMessage: true,
      messageType,
      message,
    });
    if (closeDelay && closeDelay > 0)
      setTimeout(() => this.setState({
        showMessage: false
      }), closeDelay);
  }
  
  render() {
    if (!this.props.isOfficeInitialized)
      return (
        <Stack className="page" verticalAlign="center">
          <Spinner size={SpinnerSize.large} label="Loading Office ..." />
        </Stack>
      );
    else return (
      <div>
        {!this.state.showMessage ? undefined : 
          <Stack className="page" style={{zIndex: 1000}}>
            <MessageBar messageBarType={this.state.messageType}
              isMultiline={false}
              onDismiss={() => this.setState({showMessage: false})}>
              {this.state.message}
            </MessageBar>
          </Stack>
        }
        <Stack className="page" verticalAlign="space-between"
          tokens={{ childrenGap: 2 }}>
          {
            this.props.isOfficeWordEnvironment || 
            !this.state.environmentWarningMessageBar ? undefined :
            <MessageBar messageBarType={MessageBarType.warning}
              isMultiline={false}
              title={environmentWarningMessage}
              onDismiss={() => this.setState({environmentWarningMessageBar: false})}>
              {environmentWarningMessage}
            </MessageBar>
          }
          <Header onOpenClicked={this.props.buttonOpenHandler}
            onMarkClicked={this.props.buttonMarkHandler}
            onUnmarkClicked={this.props.buttonUnmarkHandler}
            onLanguageChanged={this.onChangeLanguage}
            supportLanguageOptions={this.props.supportLanguages}
            currentLanguageId={this.props.modelLanguageId}
            commonLanguageIds={this.state.commonLanguages}
          />
          <div id={this.props.editorContainerID}></div>
          <Footer onSettingsClicked={() => this.openSettingsPanel()}
            onDownloadClicked={this.props.buttonDownloadHandler}
            onSaveClicked={this.onSaveCode}
            onInsertClicked={this.onInsertCode}
            onInsertMarkClicked={this.onInsertMarkedCode}
            isOfficeWordEnvironment={this.props.isOfficeWordEnvironment}/>
        </Stack>
        <Panel
          type={PanelType.smallFluid}
          isOpen={this.state.isSettingsOpen}
          hasCloseButton={false}
          onDismiss={this.closeSettingsPanel}
          onRenderHeader={this.onRenderSettingsPanelHeader}
          isFooterAtBottom={true}>
          <Settings ref={this.settingsPanelRef}
            tabId={this.state.settingPanelTabId}
            isOfficeWordEnvironment={this.props.isOfficeWordEnvironment}
            supportLanguages={ this.props.supportLanguages }/>
        </Panel>
      </div>
    );
  }
}