import {
  Dropdown,
  DropdownMenuItemType,
  IconButton,
  IDropdownOption,
  Stack
} from "@fluentui/react";
import * as React from "react"

export interface AppProps { 
  onOpenClicked,
  onLanguageChanged,
  onMarkClicked,
  onUnmarkClicked,
  currentLanguageId: string,
  commonLanguageIds: string[],
  supportLanguageOptions: IDropdownOption[],
}

export interface AppState { 
  languageOptions,
}

export default class Header extends React.Component<AppProps, AppState> {
  constructor(props, context) { 
    super(props, context);
    this.state = {
      languageOptions: [],
    };
  }

  static getDerivedStateFromProps(props: AppProps, _) {
    return {
      languageOptions: getLanguageOptions(
        props.currentLanguageId,
        props.commonLanguageIds,
        props.supportLanguageOptions,
      )
    };
  }

  changeLanguage = (
    event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption,
    index?: number) => {
    if (option.key != "manage")
      this.setState({
        languageOptions: getLanguageOptions(
          option.key as string,
          this.props.commonLanguageIds,
          this.props.supportLanguageOptions
        )
      });
    return this.props.onLanguageChanged(event, option, index);
  }

  render() {
    return (
      <Stack className="header" horizontal horizontalAlign="space-between">
        <Stack horizontal tokens={{ childrenGap: 5 }}>
          <IconButton title="Open File"
            onClick={this.props.onOpenClicked}
            iconProps={{ iconName: "OpenFile" }}/>
          <Dropdown placeholder="Select language"
            selectedKey={this.props.currentLanguageId}
            onChange={this.changeLanguage}
            options={this.state.languageOptions}
            dropdownWidth="auto"
            styles={{dropdown: {minWidth: 140}}} />
        </Stack>
        <Stack horizontal tokens={{ childrenGap: 5 }}>
          <IconButton title="Mark Selected Lines"
            onClick={this.props.onMarkClicked}
            iconProps={{ iconName: "SingleBookmarkSolid" }}/>
          <IconButton title="Unmark Selected Lines"
            onClick={this.props.onUnmarkClicked}
            iconProps={{ iconName: "SingleBookmark" }}/>
        </Stack>
      </Stack>
    );
  }
}

const LanguageOptionsHeaders = {
  "currentHeader": {
    key: "currentHeader", text: "Current Language",
    itemType: DropdownMenuItemType.Header
  },
  "commonHeader": {
    key: "commonHeader", text: "Commonly Used Languages",
    itemType: DropdownMenuItemType.Header,
  },
  "otherHeader": {
    key: "otherHeader", text: "Other Languages",
    itemType: DropdownMenuItemType.Header
  },
};

const LanguageOptionsDiverders: IDropdownOption[] = [
  {
    key: "divider0", text: "-",
    itemType: DropdownMenuItemType.Divider
  },
  {
    key: "divider1", text: "-",
    itemType: DropdownMenuItemType.Divider
  },
];

function getLanguageOptions(
  currentLanguageId: string,
  commonLanguageIds: string[],
  supportLanguageOptions: IDropdownOption[],
) {
  const current: IDropdownOption[] = !currentLanguageId ? [] :
    [
      LanguageOptionsHeaders["currentHeader"],
      ...supportLanguageOptions.filter(({ key }) =>
        key as string === currentLanguageId),
      LanguageOptionsDiverders[0],
    ];
  let common: IDropdownOption[] =
    supportLanguageOptions.filter(({ key }) =>
      key as string !== currentLanguageId &&
      commonLanguageIds.indexOf(key as string) !== -1);
  if (common.length > 0)
    common = [
      LanguageOptionsHeaders["commonHeader"],
      ...common,
      LanguageOptionsDiverders[1],
    ];
  const manage: IDropdownOption[] = [
    LanguageOptionsHeaders["otherHeader"],
    { key: "manage", text: "Manage Laguages" },
  ];
  return [...current, ...common, ...manage];
}