import {
  IChoiceGroupOption,
  IColorCellProps,
  IDropdownOption,
} from "@fluentui/react";

export enum LineNumberType {
  Keep = "keep",
  Reorder = "reorder",
};

export const lineNumberOptions: IChoiceGroupOption[] = [
  { key: LineNumberType.Keep, text: "Keep Original Line Number" },
  { key: LineNumberType.Reorder, text: "Reorder Line Number" },
];

export const shadingColorCellsDefault: IColorCellProps[] = [
  { id: "none", color: "#ffffff", label: "No Color" },
  { id: "white", color: "#ffffff", label: "White" },
  { id: "gray95", color: "#f3f3f3", label: "White, Darker 5%" },
  { id: "gray90", color: "#e6e6e6", label: "White, Darker 10%" },
  { id: "gray85", color: "#dadada", label: "White, Darker 15%" },
  { id: "gray80", color: "#cdcdcd", label: "White, Darker 20%" },
  { id: "gray75", color: "#c0c0c0", label: "White, Darker 25%" },
];

export const fontFamilyOptions: IDropdownOption<any>[] = [
  { key: "consolas", text: "Consolas" },
];

export const textEncodingOptions: IDropdownOption<any>[] = [
  { key: "utf8", text: "UTF-8" },
  { key: "gbk", text: "Simplified Chinese(GBK)" },
];