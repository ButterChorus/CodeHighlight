import * as monaco from "monaco-editor";

export interface SettingsOptions {
  fontSize?: string,
  fontFamily?: string,
  lineNumber?: boolean,
  lineNumberSpace?: string,
  lineNumberType?: string,
  shadingColor?: string,
  border?: boolean,
  encoding?: string,
  autoOpen?: boolean,
  commonLanguages?: string[],
  reopenFile?: boolean,
  keepMark?: boolean,
}

export interface Code {
  id: string,
  languageId: string,
  code: string,
  markRanges: monaco.Range[],
}

export class OptionsSaver {
  static defaultOptions: SettingsOptions = {
    fontSize: "12",
    fontFamily: "consolas",
    lineNumber: true,
    lineNumberSpace: "2",
    lineNumberType: "keep",
    shadingColor: "No Color",
    border: false,
    encoding: "utf8",
    autoOpen: false,
    commonLanguages: ["plaintext", "c", "java", "javascript"],
    reopenFile: true,
    keepMark: true,
  }
  private stringToValueType(key, value): any {
    if (key === "commonLanguages") {
      if (value !== undefined)
        return (value as string).split(",");
      else return [];
    }
    if (value === "false") return false;
    if (value === "true") return true;
    return value;
  }
  private valueTypeToString(key, value): string {
    if (key === "commonLanguages")
      return (value as string[]).join(",");
    if (typeof value === "number" || typeof value === "boolean")
      return (value as number|boolean).toString();
    return value;
  }
  public async setDocOptions(options: SettingsOptions,
    isOfficeWordEnvironment?: boolean) {
    if (!isOfficeWordEnvironment)
      return this.setCookieOptions(options);
    for (let key in options)
      await this.setDocProperty(
        key, this.valueTypeToString(key, options[key]));
    if (options.autoOpen)
      Office.context.document.settings.set(
        "Office.AutoShowTaskpaneWithDocument", true);
    else
      Office.context.document.settings.set(
        "Office.AutoShowTaskpaneWithDocument", false);
    Office.context.document.settings.saveAsync();
  }
  public async getDocOptions(isOfficeWordEnvironment?: boolean) {
    if (!isOfficeWordEnvironment)
      return this.getCookieOptions();
    let options: SettingsOptions = {};
    const docOptions = await this.getDocProerty();
    for (let key in OptionsSaver.defaultOptions) {
      // Notice that documnet options type could be boolean
      const doc = !docOptions[key] && docOptions[key] !== false ?
        undefined : this.stringToValueType(key, docOptions[key]);
      let cookie = this.getCookie(key);
      cookie = cookie ? this.stringToValueType(key, cookie) : undefined;
      options[key] =
        doc !== undefined ? doc :
        cookie !== undefined ? cookie :
        OptionsSaver.defaultOptions[key];
    }
    return options;
  }
  public async setCookieOptions(options: SettingsOptions) {
    for (let key in options)
      this.setCookiePermanently(
        key, this.valueTypeToString(key, options[key]));
  }
  public async getCookieOptions() {
    let options: SettingsOptions = {};
    for (let key in OptionsSaver.defaultOptions) {
      let cookie = this.getCookie(key);
      cookie = cookie ? this.stringToValueType(key, cookie) : undefined;
      options[key] =
        cookie !== undefined ? cookie :
        OptionsSaver.defaultOptions[key];
    }
    return options;
  }
  public async setDocProperty(key: string, value: string) {
    await Word.run(async (context) => {
      context.document.properties.customProperties.add(key, value);
      await context.sync();
    });
  }
  public async getDocProerty() {
    let result = {};
    await Word.run(async (context) => {
      let properties = context.document.properties.customProperties;
      properties.load("key,value");
      await context.sync();
      for (let i = 0; i < properties.items.length; i++)
        result[properties.items[i].key] = properties.items[i].value;
    });
    return result;
  }
  public setCookiePermanently(key: string, value: string) {
    let date = new Date();
    date.setFullYear(date.getFullYear() + 100)
    const expires = date.toUTCString();
    document.cookie = key + "=" + value + "; expires=" + expires + "; path=/";
  }
  public getCookie(key: string) {
    const cookies = document.cookie.split(";");
    for (let i = 0; i < cookies.length; i++) {
      const cookie = cookies[i].trim();
      if (cookie.indexOf(key + "=") == 0)
        return cookie.substring(key.length + 1, cookie.length);
    }
    return "";
  }
}

const codeXmlNamespace = "extrakit.word.codehighlight"

const codesTemplate = '<root xmlns="{{ns}}">{{codes}}</root>'
  .replace("{{ns}}", codeXmlNamespace);

const codeTemplate = `
<code xmlns="extrakit.word.codehighlight.code">
  <name xmlns="extrakit.word.codehighlight.code.name">{{name}}</name>
  <language xmlns="extrakit.word.codehighlight.code.language">{{language}}</language>
  <marks xmlns="extrakit.word.codehighlight.code.marks">{{marks}}</marks>
  <text xmlns="extrakit.word.codehighlight.code.text">{{code}}</text>
</code>
`.replace(/\n\s*/g, "");

export class CodeSaver {
  private markRangesToString(marks: monaco.Range[]) {
    let ranges = [];
    marks.forEach(range => {
      ranges.push(range.startLineNumber + ":" + range.endLineNumber);
    });
    return ranges.join(",");
  }
  private stringToMarkRanges(marks: string) {
    let ranges: monaco.Range[] = [];
    marks.split(",").forEach(range => {
      const start = Number(range.split(":")[0]);
      const end = Number(range.split(":")[1]);
      if (!isNaN(start) && !isNaN(end))
        ranges.push(new monaco.Range(start, 0, end, 0));
    });
    return ranges;
  }

  private async getXmlNodes(
    xml: Office.CustomXmlPart | Office.CustomXmlNode,
    xPath: string,
  ) {
    return new Promise<Office.CustomXmlNode[]>(
      (resolve, reject) => {
        xml.getNodesAsync(xPath, {}, result => {
          if (result.error) reject(result.error);
          else resolve(result.value);
        });
      });
  }

  private async getXmlNodeValue(xml: Office.CustomXmlNode) {
    return new Promise<string>(
      (resolve, reject) => {
        xml.getTextAsync({}, result => {
          if (result.error) reject(result.error);
          else resolve(result.value);
        });
      });
  }

  private async getCustomXmlParts(ns: string) {
    return new Promise<Office.CustomXmlPart[]>(
      (resolve, reject) => {
        Office.context.document.customXmlParts.getByNamespaceAsync(
          ns, {}, (result) => {
            if (result.error) reject(result.error);
            else resolve(result.value);
          });
      });
  }

  private async getCustomXmlNodes(ns: string, xPath: string):
    Promise<Office.CustomXmlNode[]> {
    let codeNodes = [];
    const codeXmlParts = await this.getCustomXmlParts(ns);
    if (codeXmlParts.length == 0) return [];
    await Promise.all(codeXmlParts.map(async part => {
      codeNodes.push(...await this.getXmlNodes(part, xPath));
    }));
    return codeNodes;
  }

  private async getCodeFromXmlNode(
    xml: Office.CustomXmlNode): Promise<Code|undefined> {
    const nameNodes = await this.getXmlNodes(xml, "ns2:name");
    const languageNodes = await this.getXmlNodes(xml, "ns3:language");
    const marksNodes = await this.getXmlNodes(xml, "ns4:marks");
    const textNodes = await this.getXmlNodes(xml, "ns5:text");
    if (nameNodes.length == 0 || languageNodes.length == 0 ||
      marksNodes.length == 0 || textNodes.length == 0) return undefined;
    const id = await this.getXmlNodeValue(nameNodes[0]);
    const languageId = await this.getXmlNodeValue(languageNodes[0]);
    const markRanges = this.stringToMarkRanges(
      await this.getXmlNodeValue(marksNodes[0]));
    const code = await this.getXmlNodeValue(textNodes[0]);
    return { id, languageId, markRanges, code };
  }

  private async getCodesFromXmlNodes(
    xmlNodes: Office.CustomXmlNode[]): Promise<Code[]> {
    let codes: Code[] = [];
    await Promise.all(xmlNodes.map(async node => {
      const code = await this.getCodeFromXmlNode(node);
      if (code) codes.push(code);
    }));
    return codes;
  }

  private async addCode(/*code: Code*/) {
    // TODO: Add code here
  }

  private async replaceCode(/*code: Code*/) {
    // TODO: Add code here
  }

  private async deleteCode(/*id: string*/) {
    // TODO: Add code here
  }

  public async getSavedCodes(): Promise<Code[]> { 
    const savedCodeNodes = await this.getCustomXmlNodes(
      codeXmlNamespace, '/ns0:root/ns1:code');
    return this.getCodesFromXmlNodes(savedCodeNodes);
  }

  public async getTempCodes(): Promise<Code[]> {
    const tempCodeNodes = await this.getCustomXmlNodes(codeXmlNamespace,
      '/ns0:root/ns1:code[ns2:name = "temp"]');
    return this.getCodesFromXmlNodes(tempCodeNodes);
  }

  public async saveTempCode(code: Code) {
    // TODO: Consider save multiple codes
    const codeXmlParts = await this.getCustomXmlParts(codeXmlNamespace);
    if (codeXmlParts.length > 0)
      codeXmlParts.forEach(async xmlPart => {
        await new Promise(
          (resolve, reject) => {
            xmlPart.deleteAsync({}, (result) => {
              if (result.error) reject(result.error);
              else resolve(result.value);
            });
          });
      });
    return new Promise<Office.CustomXmlPart>(
      (resolve, reject) => {
        const escapeCode = code.code
          .replace(/\&/g, "&amp;")
          .replace(/\</g, "&lt;")
          .replace(/\>/g, "&gt;")
          .replace(/\"/g, "&quot;")
          .replace(/\'/g, "&apos;")
        const tempCode = codeTemplate
          .replace("{{name}}", code.id)
          .replace("{{language}}", code.languageId)
          .replace("{{marks}}", this.markRangesToString(code.markRanges))
          .replace("{{code}}", escapeCode);
        const codeXml = codesTemplate.replace("{{codes}}", tempCode);
        Office.context.document.customXmlParts.addAsync(
          codeXml, {}, (result) => {
            if (result.error) reject(result.error);
            else resolve(result.value);
          });
      });
  }
}