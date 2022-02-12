import * as React from "react"
import {
  Checkbox,
  ChoiceGroup,
  ColorPicker,
  Dropdown,
  HoverCard,
  HoverCardType,
  IconButton,
  IDropdownOption,
  IPlainCardProps,
  Label,
  MessageBar,
  MessageBarType,
  Pivot, PivotItem,
  SearchBox,
  Stack,
  SwatchColorPicker,
  TextField,
} from "@fluentui/react";
import {
  shadingColorCellsDefault,
  fontFamilyOptions,
  lineNumberOptions,
  textEncodingOptions,
} from "./SettingsOptionsLists"
import {OptionsSaver, SettingsOptions} from "../saver"

const pivotItemStyle: React.CSSProperties = {
  margin: "5px 10px",
}

const lanugageSelectionZoneStyle: React.CSSProperties = {
  position: "absolute",
  top: 300,
  bottom: 20,
  left: 45,
  right: 0,
  overflowY: "hidden",
  overflowX: "hidden",
}

export interface AppProps {
  supportLanguages: IDropdownOption[],
  isOfficeWordEnvironment: boolean,
  tabId: string,
}

export interface AppState extends SettingsOptions {
  languageFilterText: string,
  fontSizeErrorMessage: string,
  lineNumberSpaceErrorMessage: string,
  shadingColorErrorMessage: string,
  errorState: boolean,
}

export default class Settings extends React.Component<AppProps, AppState> {
  saver: OptionsSaver;
  oldEncoding: string;
  constructor(props, context) {
    super(props, context);
    this.state = {
      ...OptionsSaver.defaultOptions,
      languageFilterText: "",
      fontSizeErrorMessage: "",
      lineNumberSpaceErrorMessage: "",
      shadingColorErrorMessage: "",
      errorState: false,
    };
    this.saver = new OptionsSaver();
  }

  componentDidMount() {
    this.loadOptions().then((_) => {
      this.oldEncoding = this.state.encoding;
    });
  }

  verifyInput = () => {
    const fontSize = this.state.fontSize;
    const lineNumberSpace = this.state.lineNumberSpace;
    const shadingColor = this.state.shadingColor;
    const fontSizeErrorMessage = /^[1-9]\d*$/.test(fontSize) ? "" : "Invalid Format";
    const lineNumberSpaceErrorMessage = /^[1-9]\d*$/.test(lineNumberSpace) ? "" : "Invalid Format";
    const shadingColorErrorMessage =
      /^(#[0-9A-Fa-f]{6})|(no color)$/.test(shadingColor.toLowerCase()) ? "" : "Invalid Format";
    const errorState = fontSizeErrorMessage ||
      lineNumberSpaceErrorMessage ||
      shadingColorErrorMessage ? true : false;
    this.setState({
      fontSizeErrorMessage,
      lineNumberSpaceErrorMessage,
      shadingColorErrorMessage,
      errorState,
    });
    return !errorState;
  }

  onFontSizeChange = (event) => {
    const value = event.target.value;
    // TODO: Verify input format
    this.setState({ fontSize: value });
  }
  onShadingColorChange = (event) => {
    const value = event.target.value;
    // TODO: Verify input format
    this.setState({ shadingColor: value });
  }
  onLineNumberSpaceChange = (event) => {
    const value = event.target.value;
    // TODO: Verify input format
    this.setState({ lineNumberSpace: value });
  }
  loadOptions = async () => {
    const options = await this.saver.getDocOptions(this.props.isOfficeWordEnvironment);
    this.setState(options);
  }
  loadOptionsFromCookie = async () => {
    const options = await this.saver.getCookieOptions();
    this.setState(options);
  }
  resetOptionsToDefaults = () => {
    this.setState(OptionsSaver.defaultOptions);
  }
  saveOptions = async () => {
    if (!this.verifyInput())
      return false;
    let options = {};
    for (let key in OptionsSaver.defaultOptions)
      // Notice that options type could be boolean
      options[key] = this.state[key] !== "" ?
        this.state[key] : OptionsSaver.defaultOptions[key];
    await this.saver.setDocOptions(options, this.props.isOfficeWordEnvironment);
    return true;
  }
  saveOptionsToCookie = async () => {
    if (!this.verifyInput()) return false;
    let options = {};
    for (let key in OptionsSaver.defaultOptions)
      options[key] = this.state[key] !== "" ?
        this.state[key] : OptionsSaver.defaultOptions[key];
    await this.saver.setCookieOptions(options);
    return true;
  }
  render() {
    const colorPickerCardProps: IPlainCardProps = {
      onRenderPlainCard: (): JSX.Element => (
        <div>
          <ColorPicker showPreview
            alphaType="none"
            onChange={(_, color) => {this.setState({shadingColor: color.str})}}
            color={this.state.shadingColor == "No Color" ?
              "#ffffff" : this.state.shadingColor} />
        </div>
      ),
    };
    return (
      <div>
        {this.state.errorState ?
          <div>
            <Stack style={{height: 2}}/>
            <MessageBar messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={() => this.setState({ errorState: false })}>
              Some input is invalid.
            </MessageBar>
            <Stack style={{height: 2}}/>
          </div>
          : undefined}
        <Pivot defaultSelectedKey={this.props.tabId}>
          <PivotItem headerText="Font" itemKey="tab1">
            <Stack style={pivotItemStyle} tokens={{ childrenGap: 2 }}>
              <TextField label="Font Size"
                value={this.state.fontSize}
                onChange={this.onFontSizeChange}
                errorMessage={this.state.fontSizeErrorMessage} />
              <Dropdown label="Font Family"
                selectedKey={this.state.fontFamily}
                placeholder="Select Font Family"
                onChange={(_, option) => {this.setState({fontFamily: option.key as string})}}
                options={fontFamilyOptions} />
              <Stack style={{height: 10}}/>
              <Checkbox label="Insert Line Number"
                onChange={(_, lineNumber) => {this.setState({lineNumber})}}
                checked={this.state.lineNumber}/>
              <ChoiceGroup options={lineNumberOptions}
                title="Line number order mode. Only worked for inserting marked lines"
                selectedKey={this.state.lineNumberType}
                onChange={(_, option) => {this.setState({lineNumberType: option.key})}}
                disabled={!this.state.lineNumber}
                style={{ margin: "5px 0 0 15px" }} />
              <TextField label="Line Number Spaces"
                value={this.state.lineNumberSpace}
                onChange={this.onLineNumberSpaceChange}
                disabled={!this.state.lineNumber}
                errorMessage={this.state.lineNumberSpaceErrorMessage}
                title="Spaces between line number and code"/>
            </Stack>
          </PivotItem>
          <PivotItem headerText="Paragraph" itemKey="tab2">
            <Stack style={pivotItemStyle} tokens={{ childrenGap: 2 }}>
              <Checkbox label="Border"
                checked={this.state.border}
                onChange={(_, border) => this.setState({border})}/>
              <Stack style={{ height: 5 }}/>
              <TextField label="Shading Color"
                onChange={this.onShadingColorChange}
                value={this.state.shadingColor}
                errorMessage={this.state.shadingColorErrorMessage} />
              <Stack horizontal verticalAlign="center">
                <SwatchColorPicker colorCells={shadingColorCellsDefault}
                  defaultSelectedId="none"
                  onChange={(_, id, color) => {
                    if (id == "none") color = "No Color";
                    this.setState({ shadingColor: color });
                  }}
                  cellShape={"square"}
                  columnCount={shadingColorCellsDefault.length} />
                <HoverCard sticky instantOpenOnClick
                  cardOpenDelay={50000000}
                  type={HoverCardType.plain}
                  plainCardProps={colorPickerCardProps}>
                  <IconButton title="More Colors"
                    iconProps={{ iconName: "More" }} />                
                </HoverCard>
              </Stack>
            </Stack>
          </PivotItem>
          <PivotItem headerText="Add-In" itemKey="tab3">
            <Stack style={pivotItemStyle} verticalAlign="space-between">
              <Stack tokens={{ childrenGap: 2 }}>
                <Checkbox label="Auto Open Taskpane"
                  title="Automatically open taskpane with document"
                  checked={this.state.autoOpen}
                  onChange={(_, checked) => this.setState({autoOpen: checked})}/>
                <Stack style={{ height: 5 }}/>
                <Dropdown label="Encoding"
                  title="Open file with encoding"
                  options={textEncodingOptions}
                  onChange={(_, option) => {this.setState({encoding: option.key as string})}}
                  selectedKey={this.state.encoding} />
                <Stack style={{ height: 5 }}/>
                <Stack horizontal tokens={{ childrenGap: 20 }}>
                  <Checkbox label="Reopen File" title="Reopen file when enconding changed"
                    checked={this.state.reopenFile}
                    onChange={(_, checked) => this.setState({reopenFile: checked})}/>
                  <Checkbox label="Keep Mark" title="Keep marks when reopen file"
                    disabled={!this.state.reopenFile}
                    checked={this.state.keepMark}
                    onChange={(_, checked) => this.setState({keepMark: checked})}/>
                </Stack>
                <Stack style={{ height: 5 }} />
                <Label>Commonly Used Code Lanugages</Label>
                <SearchBox underlined iconProps={{ iconName: "Filter" }}
                  placeholder="Filter"
                  onChange={(_, value) => this.setState({
                    languageFilterText: value !== undefined ? value : ""
                  })} />
              </Stack>
              <Stack style={lanugageSelectionZoneStyle}>
                <Stack.Item style={{ overflowY: "auto", marginRight: "-18px" }}>
                  {this.props.supportLanguages.filter((value) => {
                    const filter = this.state.languageFilterText.toLowerCase();
                    const text = value.title.toLowerCase();
                    if (filter === "" || filter === ".") return true;
                    if (filter.length === 1 && text.indexOf(filter) === 0) return true;
                    if (filter.length > 1 && text.indexOf(filter) !== -1) return true;
                    return false;
                  }).map((value) => {
                    return (
                      <Checkbox label={value.text} title={value.title} key={value.key}
                        checked={this.state.commonLanguages.indexOf(value.key as string) != -1}
                        onChange={(_, checked) => {
                          if (checked) this.setState({
                            commonLanguages: [...this.state.commonLanguages, value.key as string]
                          });
                          else this.setState({
                            commonLanguages: this.state.commonLanguages.filter(v => v != value.key)
                          });
                        }}
                        styles={{root: { margin:"2px 0" }}}/>
                    )
                  })}
                </Stack.Item>
              </Stack>
            </Stack>
          </PivotItem>
        </Pivot>
      </div>
    )
  }
}