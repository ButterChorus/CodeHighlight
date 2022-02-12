import {
  CommandBarButton,
  Label,
  MessageBarType,
  OverflowSet,
  Stack
} from "@fluentui/react";
import * as React from "react"
import Settings from "./Settings";

const messageBarCloseDelay = 1000;

export interface AppProps { 
  settingsPanelRef: React.RefObject<Settings>
  onEncodingChanged,
  onCloseSettingsPanel,
  showMessage,
  isOfficeWordEnvironment: boolean,
}

export interface AppState { 

}

export default class SettingsPanelHeader extends React.Component<AppProps, AppState> {
  render() {
    const settingsPanel = this.props.settingsPanelRef.current;
    return (
      <Stack style={{ margin: "0 30px" }} horizontal horizontalAlign="space-between">
        <Label style={{ fontSize: 18 }}>Settings</Label>
        <Stack tokens={{ childrenGap: 5 }} horizontal>
          <OverflowSet role="menubar"
            items={[
              {
                key: "reset", icon: "AppIconDefault", title: "Reset Defaults",
                onClick: () => { settingsPanel.resetOptionsToDefaults(); }
              },
              {
                key: "save", icon: "Accept", title: "Save Settings",
                onClick: () => {
                  settingsPanel.saveOptions().then(valid => {
                    if (!valid) return;
                    const { state, oldEncoding } = settingsPanel;
                    const newEncoding = state.encoding;
                    if (state.reopenFile && newEncoding != oldEncoding)
                      this.props.onEncodingChanged(newEncoding, state.keepMark);
                    this.setState({ commonLanguages: state.commonLanguages });
                    this.props.onCloseSettingsPanel();
                    this.props.showMessage(MessageBarType.success,
                      "Save settings successfully!", messageBarCloseDelay);
                  }).catch(() =>
                    this.props.showMessage(MessageBarType.error,
                      "Save settings failed!", messageBarCloseDelay));
                },
              },
              {
                key: "cancel", icon: "Cancel", title: "Cancel",
                onClick: this.props.onCloseSettingsPanel,
              },
            ]}
            overflowItems={
              !this.props.isOfficeWordEnvironment ? [] : [
                {
                  key: "saveAsDefault", name: "Save to Cookie",
                  onClick: () => { settingsPanel.saveOptionsToCookie() },
                },
                {
                  key: "loadFromDefault", name: "Load from Cookie",
                  onClick: () => { settingsPanel.loadOptionsFromCookie() },
                },
              ]}
            onRenderOverflowButton={(overflowItems) => (
              <CommandBarButton role="menuitem"
                title="More"
                menuIconProps={{ iconName: "More" }}
                menuProps={{
                  items: overflowItems,
                  isBeakVisible: false,
                  alignTargetEdge: true,
                  styles: { root: { minWidth: 100 } },
                }} />
            )}
            onRenderItem={(item) => (
              <CommandBarButton role="menuitem"
                title={item.title}
                iconProps={{ iconName: item.icon }}
                onClick={item.onClick || (() => undefined)} />
            )} />
        </Stack>
      </Stack>
    );
  }
}