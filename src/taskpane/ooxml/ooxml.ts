import { BorderTemplate, FillColorNodeTemplate, PackageTemplate, ParagraphTemplate, TextTemplate } from './tmpl'

const colors = {
  "mtk1": "#000000",  // plain text
  "mtk2": "#fffffe",
  "mtk3": "#808080",
  "mtk4": "#ff0000",
  "mtk5": "#0451a5",
  "mtk6": "#0000ff",  // keywords
  "mtk7": "#09885a",  // number
  "mtk8": "#008000",  // comments
  "mtk9": "#dd0000",
  "mtk10": "#383838",
  "mtk11": "#cd3131",
  "mtk12": "#863b00",
  "mtk13": "#af00db",
  "mtk14": "#800000",
  "mtk15": "#e00000",
  "mtk16": "#3030c0",
  "mtk17": "#666666",
  "mtk18": "#778899",
  "mtk19": "#ff00ff",
  "mtk20": "#a31515", // string
  "mtk21": "#4f76ac",
  "mtk22": "#008080",
  "mtk23": "#001188",
  "mtk24": "#4864aa",
  "lineNumbers": "#237893",
}

export interface OfficeFormatOptions {
  lineNumber?: boolean,
  maxLineNumber?: number,
  lineNumberSpace?: number,
  lineNumberColor?: string,
  fontFamily?: string,
  fontSize?: number,
  lineHeight?: number,
  shadingColor?: string,
  border?: boolean,
}

export default class OOXML {
  private formatOptions: OfficeFormatOptions;
  private pTmpl: string;
  private lines: string[] = ["<w:p/>"];
  private lineNumbers: number[] = [0];
  constructor(options?: OfficeFormatOptions) {
    this.formatOptions = { 
      lineNumber: false,
      maxLineNumber: 0,
      lineNumberSpace: 2,
      lineNumberColor: colors.lineNumbers,
      fontFamily: "Consolas",
      fontSize: 12,
      lineHeight: 320,
      shadingColor: "No Color",
      border: false,
      ...options,
    };
    this.updateParagraphTemplate();
  }

  private updateParagraphTemplate() {
    const lineNumberLength = this.formatOptions.maxLineNumber.toString().length;
    const indentation = 144 * (lineNumberLength + this.formatOptions.lineNumberSpace);
    const fillColorNode = this.formatOptions.shadingColor.toLowerCase() == "no color" ? "" :
      FillColorNodeTemplate.replace("{{fillColor}}", this.formatOptions.shadingColor);
    const borderNode = this.formatOptions.border ? BorderTemplate : "";
    this.pTmpl = ParagraphTemplate
      .replace("{{lineHeight}}", this.formatOptions.lineHeight.toString())
      .replace("{{fillColorNode}}", fillColorNode)
      .replace("{{borderNode}}", borderNode)
      .replace(/\{\{indentation\}\}/g, indentation.toString());
  }

  public setFormatOptions(options: OfficeFormatOptions) {
    this.formatOptions = { ...this.formatOptions, ...options };
    this.updateParagraphTemplate();
  }

  public addLine(htmlString: string, lineNumber: number) {
    const dom = new DOMParser().parseFromString(htmlString, "text/html");
    const textSpans = dom.querySelectorAll("span > span");
    let texts: string[] = [];
    if (this.formatOptions.lineNumber) {
      const lineNumberLength = this.formatOptions.maxLineNumber.toString().length;
      const lineNumberText = (Array(lineNumberLength).join(' ') + lineNumber.toString())
        .slice(-lineNumberLength) + Array(this.formatOptions.lineNumberSpace + 1).join(' ');
      texts.push(TextTemplate
        .replace("{{font}}", this.formatOptions.fontFamily)
        .replace("{{fontSize}}", (this.formatOptions.fontSize * 2).toString())
        .replace("{{textColor}}", this.formatOptions.lineNumberColor)
        .replace("{{text}}", lineNumberText));
    }
    textSpans.forEach(value => {
      const textColor = value.className in colors ? colors[value.className] : "#000000";
      texts.push(TextTemplate
        .replace("{{font}}", this.formatOptions.fontFamily)
        .replace("{{fontSize}}", (this.formatOptions.fontSize * 2).toString())
        .replace("{{textColor}}", textColor)
        .replace("{{text}}", value.innerHTML))
    });
    // This function may be called in an asynchronous progress,
    // so we need to reorder the lines by line number.
    for (let i = this.lines.length; i >= 0; i--) {
      if (i == 0 || lineNumber >= this.lineNumbers[i - 1]) {
        this.lines.splice(i, 0, this.pTmpl.replace("{{texts}}", texts.join('')));
        this.lineNumbers.splice(i, 0, lineNumber);
        break;
      }
    }
  }

  public removeAllLines() {
    this.lines = ["<w:p/>"];
    this.lineNumbers = [0];
  }

  public packageAllLines() {
    const pkg = PackageTemplate
      .replace("{{paragraphs}}", this.lines.join(''))
      .replace(/&nbsp;/g, " ");
    this.removeAllLines();
    return pkg;
  }
}