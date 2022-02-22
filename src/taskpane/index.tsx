import * as monaco from "monaco-editor";
import * as React from "react"
import * as ReactDOM from "react-dom"
import { AppContainer } from "react-hot-loader";
import { initializeIcons } from '@fluentui/font-icons-mdl2'
import { IDropdownOption } from "@fluentui/react";
import App from "./components/App";
import Editor from "./editor"
import OOXML, { OfficeFormatOptions } from "./ooxml/ooxml"
import { CodeSaver, OptionsSaver, SettingsOptions } from "./saver";

let isOfficeInitialized = false;
let isOfficeWordEnvironment = false;
let editor: Editor;
let ooxml: OOXML;
let supportLanguages: IDropdownOption[] = [];

const editorContainerID = "editor-container";
const markClassName = "mark";

const render = Component => new Promise<void>(
  (resolve, _) => {
    supportLanguages = Editor.getSupportLanguages().map(
      ({ id, aliases, extensions }) => {
        const ext0 = extensions[0] ? "(" + extensions[0] + ")" : "";
        const exts = extensions[0] ? "(" + extensions.join(", ") + ")" : "";
        return {
          key: id,
          text: aliases[0] + " " + ext0,
          title: aliases.join(", ") + " " + exts,
        }
      }
    );
    ReactDOM.render(
      <AppContainer>
        <div>
          <Component
            isOfficeInitialized={isOfficeInitialized}
            isOfficeWordEnvironment={isOfficeWordEnvironment}
            editorContainerID={editorContainerID}
            supportLanguages={supportLanguages}
            modelLanguageId={editor ? editor.getLanguageId() : ""}
            buttonOpenHandler={buttonOpenClicked}
            buttonMarkHandler={buttonMarkClicked}
            buttonUnmarkHandler={buttonUnmarkClicked}
            buttonDownloadHandler={buttonDownloadClicked}
            buttonSaveHandler={buttonSaveClicked}
            buttonInsertHandler={buttonInsertClicked}
            buttonInsertMarkedHandler={buttonInsertMarkClicked}
            languageChangeHandler={lanugageChanged}
            encodingChangeHandler={encodingChanged}
          />
        </div>
      </AppContainer >,
      document.getElementById("root"),
      resolve,
    )
  }
);

initializeIcons("./fonts/");

// hide spinner
document.getElementById("loader").style.display = "none";

Promise.all([
  render(App),
  Office.onReady().then(info => {
    if (info.host === Office.HostType.Word)
      isOfficeWordEnvironment = true;
  }),
]).then(async () => {
  isOfficeInitialized = true;
  await render(App);
  let editorDOM = document.getElementById(editorContainerID);
  editor = new Editor(
    editorDOM,
    { markClassName },
  );
  addEventListeners(editorDOM);
  (new CodeSaver()).getTempCodes().then(codes => {
    if (codes.length > 0) {
      editor.openSavedCode(
        codes[0].code,
        codes[0].languageId,
        codes[0].markRanges,
      );
      render(App);
    }
  })
  ooxml = new OOXML();
}).catch(reason => console.log(reason));

if ((module as any).hot) {
  (module as any).hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    render(NextApp);
  });
}

async function buttonSaveClicked() {
  if (!isOfficeWordEnvironment) return;
  const code = {
    id: "temp",
    languageId: editor.getLanguageId(),
    code: editor.getValue(),
    markRanges: editor.getMarkedRanges(),
  };
  await (new CodeSaver()).saveTempCode(code);
  return;
}

function addEventListeners(editorDOM: HTMLElement) {
  editorDOM.addEventListener("drop", (event) => {
    event.preventDefault();
    event.stopPropagation();
    let file = event.dataTransfer.files[0];
    (new OptionsSaver()).getDocOptions(isOfficeWordEnvironment)
      .then((options) => {
      editor.open(file, options.encoding);
      render(App);
    }).catch(reason => console.log(reason));
  });
  editorDOM.addEventListener("dragenter", (event) => {
    event.preventDefault();
    event.stopPropagation();
    return false;
  });
  editorDOM.addEventListener("dragover", (event) => {
    event.preventDefault();
    event.stopPropagation();
    return false;
  });
  editorDOM.addEventListener("dragleave", (event) => {
    event.preventDefault();
    event.stopPropagation();
    return false;
  });
}

function buttonOpenClicked() {
  (new OptionsSaver()).getDocOptions(isOfficeWordEnvironment)
    .then((options) => {
    editor.openFile(options.encoding);
    render(App);
  }).catch(reason => console.log(reason));
}

function buttonMarkClicked() {
  editor.markSelectLines();
}

function buttonUnmarkClicked() {
  editor.unmarkSelectLines();
}

function buttonDownloadClicked() {
  editor.downloaad();
}

function setOfficeFormatOptions(maxLineNumber: number, settings: SettingsOptions) {
  const fontSize = Number(settings.fontSize) || 12;
  const lineNumberSpace = Number(settings.lineNumberSpace) || 2;
  const options: OfficeFormatOptions = {
    lineNumber: settings.lineNumber,
    lineNumberSpace,
    maxLineNumber,
    fontFamily: settings.fontFamily,
    fontSize,
    lineHeight: Math.round(fontSize * 1.33) * 20,
    shadingColor: settings.shadingColor,
    border: settings.border,
  };
  ooxml.setFormatOptions(options);
  if (editor.getLanguageId() === "plaintext")
    ooxml.setFormatOptions({ lineNumberColor: "#000000" });
}

async function buttonInsertClicked() {
  if (!isOfficeWordEnvironment) return;
  const lines = editor.getAllLines();
  const settings = await (new OptionsSaver())
    .getDocOptions(isOfficeWordEnvironment);
  setOfficeFormatOptions(lines.length + 1, settings);
  return Word.run(async (context) => {
    await Promise.all(lines.map(async (value, index) => {
      const htmlString = await monaco.editor.colorize(
        value, editor.getLanguageId(), {});
      ooxml.addLine(htmlString, index + 1);
    }));
    const pkg = ooxml.packageAllLines();
    await new Promise<void>(
      (resolve, reject) => {
        Office.context.document.setSelectedDataAsync(
          pkg, { coercionType: 'ooxml' },
          result => {
            if (result.error) reject(result.error);
            else resolve(result.value)
          });
      });
    await context.sync();
  });
}

async function buttonInsertMarkClicked(reorderLineNumber?: boolean) {
  if (!isOfficeWordEnvironment) return;
  const markedLinesNumber = editor.getMarkedLinesNumber();
  const lines = editor.getAllLines();
  const settings = await (new OptionsSaver())
    .getDocOptions(isOfficeWordEnvironment);
  setOfficeFormatOptions(lines.length + 1, settings);
  let i = 0;
  return Word.run(async (context) => {
    await Promise.all(markedLinesNumber.map(async (value, index) => {
      if (!value) return;
      i += 1;
      const lineNumber = reorderLineNumber ? i : index + 1;
      const htmlString = await monaco.editor.colorize(
        lines[index], editor.getLanguageId(), {});
      ooxml.addLine(htmlString, lineNumber);
    }))
    if (i > 0) {
      const pkg = ooxml.packageAllLines();
      await new Promise<void>(
        (resolve, reject) => {
          Office.context.document.setSelectedDataAsync(
            pkg, { coercionType: 'ooxml' },
            result => {
              if (result.error) reject(result.error);
              else resolve(result.value)
            });
        });
    }
    await context.sync();
  });
}

function lanugageChanged(
  _: React.FormEvent<HTMLDivElement>,
  option: IDropdownOption,
  __: number) {
  editor.changeLanguage(option.key as string);
  render(App);
}

function encodingChanged(encoding: string, keepMark: boolean) {
  editor.reopenWithEncoding(encoding, keepMark);
}
