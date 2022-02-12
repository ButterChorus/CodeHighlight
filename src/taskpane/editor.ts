import * as monaco from "monaco-editor";

export interface EditorOptions {
  markClassName?: string,
  textEncoding?: string,
}

export default class Editor implements EditorOptions {
  private editor: monaco.editor.IStandaloneCodeEditor;
  private fileInput: HTMLInputElement = document.createElement("input");
  private fileDownload: HTMLAnchorElement = document.createElement("a");
  private decorations: string[] = [];
  private languageId: string = "";
  private file: File = undefined;
  markClassName: string;
  textEncoding: string;

  constructor(domElement: HTMLElement, options: EditorOptions, listeners?) {
    this.editor = monaco.editor.create(domElement,
      {
        language: this.languageId,
        automaticLayout: true,
        contextmenu: false,
        minimap: { renderCharacters: false },
      });
    if (listeners) {
      if (listeners.onFocus) this.editor.onDidFocusEditorWidget(listeners.onFocus);
      if (listeners.onBlur) this.editor.onDidBlurEditorWidget(listeners.onBlur);
      // TODO: Add more event listeners here
    }
    this.fileInput.type = "file";
    this.markClassName = options.markClassName? options.markClassName: "";
    this.textEncoding = options.textEncoding ? options.textEncoding : "utf8";
  }

  public setOptions(options: EditorOptions) {
    for (let key in options) this[key] = options[key];
  }

  public getAllLines() {
    return this.editor.getModel().getLinesContent();
  }

  public getValue() {
    return this.editor.getValue();
  }

  public getMarkedRanges() {
    let markedRanges: monaco.Range[] = [];
    this.decorations.forEach((id) => {
      let range = this.editor.getModel().getDecorationRange(id);
      markedRanges.push(range);
    });
    return markedRanges;
  }

  public getMarkedLinesNumber() {
    const linesCount = this.editor.getModel().getLineCount();
    let markedLines = new Array(linesCount);
    this.getMarkedRanges().forEach((range) => {
      for (let i = range.startLineNumber; i <= range.endLineNumber; i++)
        markedLines[i - 1] = true;
    })
    return markedLines;
  }

  public getLanguageId() {
    return this.languageId !== "" ? this.languageId : "plaintext";
  }

  public static getSupportLanguages() {
    return monaco.languages.getLanguages();
  }

  public downloaad() {
    const content = this.editor.getValue();
    // IE
    if ((window.navigator as any).msSaveBlob)
      (window.navigator as any).msSaveBlob(new Blob([content]));
    // Explorers Support Data URI
    else if (undefined !== this.fileDownload.download) {
      this.fileDownload.href = "data:," + content;
      this.fileDownload.click();
    }
  }

  public openFile(encoding?: string) {
    this.fileInput.click();
    if (this.fileInput.files.length > 0) {
      const file = this.fileInput.files[0];
      this.open(file, encoding);
    }
  }

  public open(file: File, encoding?: string) {
    const reader = new FileReader();
    reader.readAsText(file, encoding ? encoding : this.textEncoding);
    reader.onload = event => {
      this.editor.setValue(event.target.result as string);
    };
    this.file = file;
    // Change model language.
    const ext = file.name.split(".").pop();
    const languages = Editor.getSupportLanguages().filter(value => 
      value.extensions.indexOf("." + ext) != -1
    );
    const languageId = languages.length > 0 ?
      languages[0].id : "plaintext";
    this.changeLanguage(languageId);
  }

  public reopenWithEncoding(encoding: string, keepMark?: boolean) {
    if (this.file) {
      let markedRanges: monaco.Range[] = [];
      if (keepMark)
        markedRanges = this.getMarkedRanges();
      this.decorations = [];
      const reader = new FileReader();
      reader.readAsText(this.file, encoding);
      reader.onload = event => {
        this.editor.setValue(event.target.result as string);
        if (keepMark)
          this.markRanges([], this.markClassName, markedRanges);
      };
    }
  }

  public openSavedCode(
    code: string,
    languageId?: string,
    markedRanges?: monaco.Range[]) {
    this.editor.setValue(code);
    this.decorations = [];
    this.file = undefined;
    if (languageId) this.changeLanguage(languageId);
    if (markedRanges)
      this.markRanges([], this.markClassName, markedRanges);
  }

  public changeLanguage(languageId: string) {
    monaco.editor.setModelLanguage(this.editor.getModel(), languageId);
    this.languageId = languageId;
  }

  public markSelectLines() {
    this.markContactLines(true);
  }

  public unmarkSelectLines() {
    this.markContactLines(false);
  }

  private markContactLines(mark: boolean) {
    const selection = this.editor.getSelection();
    let start = selection.startLineNumber;
    let end = selection.endLineNumber;
    const cd = this.editor.getModel().getDecorationsInRange(
      new monaco.Range(start - 1, 0, end + 1, 0),
    ).filter(({options}) => options.className == this.markClassName);
    start = Math.min(start, ...cd.map(({range}) => range.startLineNumber));
    end = Math.max(end, ...cd.map(({range}) => range.endLineNumber));
    cd.forEach(value =>
    { this.decorations.splice(this.decorations.indexOf(value.id), 1) });

    if (mark) 
      this.decorations.push(
        ...this.markLines(
          cd.map(({id}) => id), this.markClassName, start, end));
    else {
      this.markLines(cd.map(({id}) => id), "", start, end);
      if (start < selection.startLineNumber)
        this.decorations.push(
          ...this.markLines(
            [], this.markClassName, start, selection.startLineNumber - 1
          ));
      if (selection.endLineNumber < end)
        this.decorations.push(
          ...this.markLines(
            [], this.markClassName, selection.endLineNumber + 1, end
          ));
    }
  }

  private markLines(
    oldDecorations: string[],
    className: string,
    start: number,
    end: number) {
    return this.editor.deltaDecorations(
      oldDecorations,
      [{
        range: new monaco.Range(start, 0, end, 0),
        options: {
          className,
          isWholeLine: true
        },
      }]);
  }

  private markRanges(
    oldDecorations: string[],
    className: string,
    markedRanges: monaco.Range[]) {
    markedRanges.forEach((range) => {
      this.decorations.push(
        ...this.markLines(
          oldDecorations,
          className,
          range.startLineNumber,
          range.endLineNumber
        )
      );
    })
  }
}