"use strict";(self.webpackChunkoffice_addin_taskpane_react=self.webpackChunkoffice_addin_taskpane_react||[]).push([[1966],{91966:function(e,n,t){t.r(n),t.d(n,{setupMode:function(){return _e}});var r,i,o,a,u,s,c,d,f,l,g,h,p,m,v,_,b,k,y,w=function(){function e(e){var n=this;this._defaults=e,this._worker=null,this._idleCheckInterval=setInterval((function(){return n._checkIfIdle()}),3e4),this._lastUsedTime=0,this._configChangeListener=this._defaults.onDidChange((function(){return n._stopWorker()}))}return e.prototype._stopWorker=function(){this._worker&&(this._worker.dispose(),this._worker=null),this._client=null},e.prototype.dispose=function(){clearInterval(this._idleCheckInterval),this._configChangeListener.dispose(),this._stopWorker()},e.prototype._checkIfIdle=function(){this._worker&&Date.now()-this._lastUsedTime>12e4&&this._stopWorker()},e.prototype._getClient=function(){return this._lastUsedTime=Date.now(),this._client||(this._worker=monaco.editor.createWebWorker({moduleId:"vs/language/html/htmlWorker",createData:{languageSettings:this._defaults.options,languageId:this._defaults.languageId},label:this._defaults.languageId}),this._client=this._worker.getProxy()),this._client},e.prototype.getLanguageServiceWorker=function(){for(var e,n=this,t=[],r=0;r<arguments.length;r++)t[r]=arguments[r];return this._getClient().then((function(n){e=n})).then((function(e){return n._worker.withSyncedResources(t)})).then((function(n){return e}))},e}();!function(e){e.create=function(e,n){return{line:e,character:n}},e.is=function(e){var n=e;return J.objectLiteral(n)&&J.number(n.line)&&J.number(n.character)}}(r||(r={})),function(e){e.create=function(e,n,t,i){if(J.number(e)&&J.number(n)&&J.number(t)&&J.number(i))return{start:r.create(e,n),end:r.create(t,i)};if(r.is(e)&&r.is(n))return{start:e,end:n};throw new Error("Range#create called with invalid arguments["+e+", "+n+", "+t+", "+i+"]")},e.is=function(e){var n=e;return J.objectLiteral(n)&&r.is(n.start)&&r.is(n.end)}}(i||(i={})),function(e){e.create=function(e,n){return{uri:e,range:n}},e.is=function(e){var n=e;return J.defined(n)&&i.is(n.range)&&(J.string(n.uri)||J.undefined(n.uri))}}(o||(o={})),function(e){e.create=function(e,n,t,r){return{targetUri:e,targetRange:n,targetSelectionRange:t,originSelectionRange:r}},e.is=function(e){var n=e;return J.defined(n)&&i.is(n.targetRange)&&J.string(n.targetUri)&&(i.is(n.targetSelectionRange)||J.undefined(n.targetSelectionRange))&&(i.is(n.originSelectionRange)||J.undefined(n.originSelectionRange))}}(a||(a={})),function(e){e.create=function(e,n,t,r){return{red:e,green:n,blue:t,alpha:r}},e.is=function(e){var n=e;return J.number(n.red)&&J.number(n.green)&&J.number(n.blue)&&J.number(n.alpha)}}(u||(u={})),function(e){e.create=function(e,n){return{range:e,color:n}},e.is=function(e){var n=e;return i.is(n.range)&&u.is(n.color)}}(s||(s={})),function(e){e.create=function(e,n,t){return{label:e,textEdit:n,additionalTextEdits:t}},e.is=function(e){var n=e;return J.string(n.label)&&(J.undefined(n.textEdit)||m.is(n))&&(J.undefined(n.additionalTextEdits)||J.typedArray(n.additionalTextEdits,m.is))}}(c||(c={})),function(e){e.Comment="comment",e.Imports="imports",e.Region="region"}(d||(d={})),function(e){e.create=function(e,n,t,r,i){var o={startLine:e,endLine:n};return J.defined(t)&&(o.startCharacter=t),J.defined(r)&&(o.endCharacter=r),J.defined(i)&&(o.kind=i),o},e.is=function(e){var n=e;return J.number(n.startLine)&&J.number(n.startLine)&&(J.undefined(n.startCharacter)||J.number(n.startCharacter))&&(J.undefined(n.endCharacter)||J.number(n.endCharacter))&&(J.undefined(n.kind)||J.string(n.kind))}}(f||(f={})),function(e){e.create=function(e,n){return{location:e,message:n}},e.is=function(e){var n=e;return J.defined(n)&&o.is(n.location)&&J.string(n.message)}}(l||(l={})),function(e){e.Error=1,e.Warning=2,e.Information=3,e.Hint=4}(g||(g={})),function(e){e.create=function(e,n,t,r,i,o){var a={range:e,message:n};return J.defined(t)&&(a.severity=t),J.defined(r)&&(a.code=r),J.defined(i)&&(a.source=i),J.defined(o)&&(a.relatedInformation=o),a},e.is=function(e){var n=e;return J.defined(n)&&i.is(n.range)&&J.string(n.message)&&(J.number(n.severity)||J.undefined(n.severity))&&(J.number(n.code)||J.string(n.code)||J.undefined(n.code))&&(J.string(n.source)||J.undefined(n.source))&&(J.undefined(n.relatedInformation)||J.typedArray(n.relatedInformation,l.is))}}(h||(h={})),function(e){e.create=function(e,n){for(var t=[],r=2;r<arguments.length;r++)t[r-2]=arguments[r];var i={title:e,command:n};return J.defined(t)&&t.length>0&&(i.arguments=t),i},e.is=function(e){var n=e;return J.defined(n)&&J.string(n.title)&&J.string(n.command)}}(p||(p={})),function(e){e.replace=function(e,n){return{range:e,newText:n}},e.insert=function(e,n){return{range:{start:e,end:e},newText:n}},e.del=function(e){return{range:e,newText:""}},e.is=function(e){var n=e;return J.objectLiteral(n)&&J.string(n.newText)&&i.is(n.range)}}(m||(m={})),function(e){e.create=function(e,n){return{textDocument:e,edits:n}},e.is=function(e){var n=e;return J.defined(n)&&x.is(n.textDocument)&&Array.isArray(n.edits)}}(v||(v={})),function(e){e.create=function(e,n){var t={kind:"create",uri:e};return void 0===n||void 0===n.overwrite&&void 0===n.ignoreIfExists||(t.options=n),t},e.is=function(e){var n=e;return n&&"create"===n.kind&&J.string(n.uri)&&(void 0===n.options||(void 0===n.options.overwrite||J.boolean(n.options.overwrite))&&(void 0===n.options.ignoreIfExists||J.boolean(n.options.ignoreIfExists)))}}(_||(_={})),function(e){e.create=function(e,n,t){var r={kind:"rename",oldUri:e,newUri:n};return void 0===t||void 0===t.overwrite&&void 0===t.ignoreIfExists||(r.options=t),r},e.is=function(e){var n=e;return n&&"rename"===n.kind&&J.string(n.oldUri)&&J.string(n.newUri)&&(void 0===n.options||(void 0===n.options.overwrite||J.boolean(n.options.overwrite))&&(void 0===n.options.ignoreIfExists||J.boolean(n.options.ignoreIfExists)))}}(b||(b={})),function(e){e.create=function(e,n){var t={kind:"delete",uri:e};return void 0===n||void 0===n.recursive&&void 0===n.ignoreIfNotExists||(t.options=n),t},e.is=function(e){var n=e;return n&&"delete"===n.kind&&J.string(n.uri)&&(void 0===n.options||(void 0===n.options.recursive||J.boolean(n.options.recursive))&&(void 0===n.options.ignoreIfNotExists||J.boolean(n.options.ignoreIfNotExists)))}}(k||(k={})),function(e){e.is=function(e){var n=e;return n&&(void 0!==n.changes||void 0!==n.documentChanges)&&(void 0===n.documentChanges||n.documentChanges.every((function(e){return J.string(e.kind)?_.is(e)||b.is(e)||k.is(e):v.is(e)})))}}(y||(y={}));var C,x,E,I,S,T,M,R,F,P,A,D,L,O,j,W,N,U=function(){function e(e){this.edits=e}return e.prototype.insert=function(e,n){this.edits.push(m.insert(e,n))},e.prototype.replace=function(e,n){this.edits.push(m.replace(e,n))},e.prototype.delete=function(e){this.edits.push(m.del(e))},e.prototype.add=function(e){this.edits.push(e)},e.prototype.all=function(){return this.edits},e.prototype.clear=function(){this.edits.splice(0,this.edits.length)},e}();!function(){function e(e){var n=this;this._textEditChanges=Object.create(null),e&&(this._workspaceEdit=e,e.documentChanges?e.documentChanges.forEach((function(e){if(v.is(e)){var t=new U(e.edits);n._textEditChanges[e.textDocument.uri]=t}})):e.changes&&Object.keys(e.changes).forEach((function(t){var r=new U(e.changes[t]);n._textEditChanges[t]=r})))}Object.defineProperty(e.prototype,"edit",{get:function(){return this._workspaceEdit},enumerable:!0,configurable:!0}),e.prototype.getTextEditChange=function(e){if(x.is(e)){if(this._workspaceEdit||(this._workspaceEdit={documentChanges:[]}),!this._workspaceEdit.documentChanges)throw new Error("Workspace edit is not configured for document changes.");var n=e;if(!(r=this._textEditChanges[n.uri])){var t={textDocument:n,edits:i=[]};this._workspaceEdit.documentChanges.push(t),r=new U(i),this._textEditChanges[n.uri]=r}return r}if(this._workspaceEdit||(this._workspaceEdit={changes:Object.create(null)}),!this._workspaceEdit.changes)throw new Error("Workspace edit is not configured for normal text edit changes.");var r;if(!(r=this._textEditChanges[e])){var i=[];this._workspaceEdit.changes[e]=i,r=new U(i),this._textEditChanges[e]=r}return r},e.prototype.createFile=function(e,n){this.checkDocumentChanges(),this._workspaceEdit.documentChanges.push(_.create(e,n))},e.prototype.renameFile=function(e,n,t){this.checkDocumentChanges(),this._workspaceEdit.documentChanges.push(b.create(e,n,t))},e.prototype.deleteFile=function(e,n){this.checkDocumentChanges(),this._workspaceEdit.documentChanges.push(k.create(e,n))},e.prototype.checkDocumentChanges=function(){if(!this._workspaceEdit||!this._workspaceEdit.documentChanges)throw new Error("Workspace edit is not configured for document changes.")}}(),function(e){e.create=function(e){return{uri:e}},e.is=function(e){var n=e;return J.defined(n)&&J.string(n.uri)}}(C||(C={})),function(e){e.create=function(e,n){return{uri:e,version:n}},e.is=function(e){var n=e;return J.defined(n)&&J.string(n.uri)&&(null===n.version||J.number(n.version))}}(x||(x={})),function(e){e.create=function(e,n,t,r){return{uri:e,languageId:n,version:t,text:r}},e.is=function(e){var n=e;return J.defined(n)&&J.string(n.uri)&&J.string(n.languageId)&&J.number(n.version)&&J.string(n.text)}}(E||(E={})),function(e){e.PlainText="plaintext",e.Markdown="markdown"}(I||(I={})),function(e){e.is=function(n){var t=n;return t===e.PlainText||t===e.Markdown}}(I||(I={})),function(e){e.is=function(e){var n=e;return J.objectLiteral(e)&&I.is(n.kind)&&J.string(n.value)}}(S||(S={})),function(e){e.Text=1,e.Method=2,e.Function=3,e.Constructor=4,e.Field=5,e.Variable=6,e.Class=7,e.Interface=8,e.Module=9,e.Property=10,e.Unit=11,e.Value=12,e.Enum=13,e.Keyword=14,e.Snippet=15,e.Color=16,e.File=17,e.Reference=18,e.Folder=19,e.EnumMember=20,e.Constant=21,e.Struct=22,e.Event=23,e.Operator=24,e.TypeParameter=25}(T||(T={})),function(e){e.PlainText=1,e.Snippet=2}(M||(M={})),function(e){e.create=function(e){return{label:e}}}(R||(R={})),function(e){e.create=function(e,n){return{items:e||[],isIncomplete:!!n}}}(F||(F={})),function(e){e.fromPlainText=function(e){return e.replace(/[\\`*_{}[\]()#+\-.!]/g,"\\$&")},e.is=function(e){var n=e;return J.string(n)||J.objectLiteral(n)&&J.string(n.language)&&J.string(n.value)}}(P||(P={})),function(e){e.is=function(e){var n=e;return!!n&&J.objectLiteral(n)&&(S.is(n.contents)||P.is(n.contents)||J.typedArray(n.contents,P.is))&&(void 0===e.range||i.is(e.range))}}(A||(A={})),function(e){e.create=function(e,n){return n?{label:e,documentation:n}:{label:e}}}(D||(D={})),function(e){e.create=function(e,n){for(var t=[],r=2;r<arguments.length;r++)t[r-2]=arguments[r];var i={label:e};return J.defined(n)&&(i.documentation=n),J.defined(t)?i.parameters=t:i.parameters=[],i}}(L||(L={})),function(e){e.Text=1,e.Read=2,e.Write=3}(O||(O={})),function(e){e.create=function(e,n){var t={range:e};return J.number(n)&&(t.kind=n),t}}(j||(j={})),function(e){e.File=1,e.Module=2,e.Namespace=3,e.Package=4,e.Class=5,e.Method=6,e.Property=7,e.Field=8,e.Constructor=9,e.Enum=10,e.Interface=11,e.Function=12,e.Variable=13,e.Constant=14,e.String=15,e.Number=16,e.Boolean=17,e.Array=18,e.Object=19,e.Key=20,e.Null=21,e.EnumMember=22,e.Struct=23,e.Event=24,e.Operator=25,e.TypeParameter=26}(W||(W={})),function(e){e.create=function(e,n,t,r,i){var o={name:e,kind:n,location:{uri:r,range:t}};return i&&(o.containerName=i),o}}(N||(N={}));var V,H,K,z,B,$=function(){};!function(e){e.create=function(e,n,t,r,i,o){var a={name:e,detail:n,kind:t,range:r,selectionRange:i};return void 0!==o&&(a.children=o),a},e.is=function(e){var n=e;return n&&J.string(n.name)&&J.number(n.kind)&&i.is(n.range)&&i.is(n.selectionRange)&&(void 0===n.detail||J.string(n.detail))&&(void 0===n.deprecated||J.boolean(n.deprecated))&&(void 0===n.children||Array.isArray(n.children))}}($||($={})),function(e){e.QuickFix="quickfix",e.Refactor="refactor",e.RefactorExtract="refactor.extract",e.RefactorInline="refactor.inline",e.RefactorRewrite="refactor.rewrite",e.Source="source",e.SourceOrganizeImports="source.organizeImports"}(V||(V={})),function(e){e.create=function(e,n){var t={diagnostics:e};return null!=n&&(t.only=n),t},e.is=function(e){var n=e;return J.defined(n)&&J.typedArray(n.diagnostics,h.is)&&(void 0===n.only||J.typedArray(n.only,J.string))}}(H||(H={})),function(e){e.create=function(e,n,t){var r={title:e};return p.is(n)?r.command=n:r.edit=n,void 0!==t&&(r.kind=t),r},e.is=function(e){var n=e;return n&&J.string(n.title)&&(void 0===n.diagnostics||J.typedArray(n.diagnostics,h.is))&&(void 0===n.kind||J.string(n.kind))&&(void 0!==n.edit||void 0!==n.command)&&(void 0===n.command||p.is(n.command))&&(void 0===n.edit||y.is(n.edit))}}(K||(K={})),function(e){e.create=function(e,n){var t={range:e};return J.defined(n)&&(t.data=n),t},e.is=function(e){var n=e;return J.defined(n)&&i.is(n.range)&&(J.undefined(n.command)||p.is(n.command))}}(z||(z={})),function(e){e.create=function(e,n){return{tabSize:e,insertSpaces:n}},e.is=function(e){var n=e;return J.defined(n)&&J.number(n.tabSize)&&J.boolean(n.insertSpaces)}}(B||(B={}));var q,Q,G=function(){};!function(e){e.create=function(e,n,t){return{range:e,target:n,data:t}},e.is=function(e){var n=e;return J.defined(n)&&i.is(n.range)&&(J.undefined(n.target)||J.string(n.target))}}(G||(G={})),function(e){function n(e,t){if(e.length<=1)return e;var r=e.length/2|0,i=e.slice(0,r),o=e.slice(r);n(i,t),n(o,t);for(var a=0,u=0,s=0;a<i.length&&u<o.length;){var c=t(i[a],o[u]);e[s++]=c<=0?i[a++]:o[u++]}for(;a<i.length;)e[s++]=i[a++];for(;u<o.length;)e[s++]=o[u++];return e}e.create=function(e,n,t,r){return new X(e,n,t,r)},e.is=function(e){var n=e;return!!(J.defined(n)&&J.string(n.uri)&&(J.undefined(n.languageId)||J.string(n.languageId))&&J.number(n.lineCount)&&J.func(n.getText)&&J.func(n.positionAt)&&J.func(n.offsetAt))},e.applyEdits=function(e,t){for(var r=e.getText(),i=n(t,(function(e,n){var t=e.range.start.line-n.range.start.line;return 0===t?e.range.start.character-n.range.start.character:t})),o=r.length,a=i.length-1;a>=0;a--){var u=i[a],s=e.offsetAt(u.range.start),c=e.offsetAt(u.range.end);if(!(c<=o))throw new Error("Overlapping edit");r=r.substring(0,s)+u.newText+r.substring(c,r.length),o=s}return r}}(q||(q={})),function(e){e.Manual=1,e.AfterDelay=2,e.FocusOut=3}(Q||(Q={}));var J,X=function(){function e(e,n,t,r){this._uri=e,this._languageId=n,this._version=t,this._content=r,this._lineOffsets=null}return Object.defineProperty(e.prototype,"uri",{get:function(){return this._uri},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"languageId",{get:function(){return this._languageId},enumerable:!0,configurable:!0}),Object.defineProperty(e.prototype,"version",{get:function(){return this._version},enumerable:!0,configurable:!0}),e.prototype.getText=function(e){if(e){var n=this.offsetAt(e.start),t=this.offsetAt(e.end);return this._content.substring(n,t)}return this._content},e.prototype.update=function(e,n){this._content=e.text,this._version=n,this._lineOffsets=null},e.prototype.getLineOffsets=function(){if(null===this._lineOffsets){for(var e=[],n=this._content,t=!0,r=0;r<n.length;r++){t&&(e.push(r),t=!1);var i=n.charAt(r);t="\r"===i||"\n"===i,"\r"===i&&r+1<n.length&&"\n"===n.charAt(r+1)&&r++}t&&n.length>0&&e.push(n.length),this._lineOffsets=e}return this._lineOffsets},e.prototype.positionAt=function(e){e=Math.max(Math.min(e,this._content.length),0);var n=this.getLineOffsets(),t=0,i=n.length;if(0===i)return r.create(0,e);for(;t<i;){var o=Math.floor((t+i)/2);n[o]>e?i=o:t=o+1}var a=t-1;return r.create(a,e-n[a])},e.prototype.offsetAt=function(e){var n=this.getLineOffsets();if(e.line>=n.length)return this._content.length;if(e.line<0)return 0;var t=n[e.line],r=e.line+1<n.length?n[e.line+1]:this._content.length;return Math.max(Math.min(t+e.character,r),t)},Object.defineProperty(e.prototype,"lineCount",{get:function(){return this.getLineOffsets().length},enumerable:!0,configurable:!0}),e}();!function(e){var n=Object.prototype.toString;e.defined=function(e){return void 0!==e},e.undefined=function(e){return void 0===e},e.boolean=function(e){return!0===e||!1===e},e.string=function(e){return"[object String]"===n.call(e)},e.number=function(e){return"[object Number]"===n.call(e)},e.func=function(e){return"[object Function]"===n.call(e)},e.objectLiteral=function(e){return null!==e&&"object"==typeof e},e.typedArray=function(e,n){return Array.isArray(e)&&e.every(n)}}(J||(J={}));var Y=monaco.Range,Z=function(){function e(e,n,t){var r=this;this._languageId=e,this._worker=n,this._disposables=[],this._listener=Object.create(null);var i=function(e){var n,t=e.getModeId();t===r._languageId&&(r._listener[e.uri.toString()]=e.onDidChangeContent((function(){clearTimeout(n),n=setTimeout((function(){return r._doValidate(e.uri,t)}),500)})),r._doValidate(e.uri,t))},o=function(e){monaco.editor.setModelMarkers(e,r._languageId,[]);var n=e.uri.toString(),t=r._listener[n];t&&(t.dispose(),delete r._listener[n])};this._disposables.push(monaco.editor.onDidCreateModel(i)),this._disposables.push(monaco.editor.onWillDisposeModel((function(e){o(e)}))),this._disposables.push(monaco.editor.onDidChangeModelLanguage((function(e){o(e.model),i(e.model)}))),this._disposables.push(t.onDidChange((function(e){monaco.editor.getModels().forEach((function(e){e.getModeId()===r._languageId&&(o(e),i(e))}))}))),this._disposables.push({dispose:function(){for(var e in r._listener)r._listener[e].dispose()}}),monaco.editor.getModels().forEach(i)}return e.prototype.dispose=function(){this._disposables.forEach((function(e){return e&&e.dispose()})),this._disposables=[]},e.prototype._doValidate=function(e,n){this._worker(e).then((function(t){return t.doValidation(e.toString()).then((function(t){var r=t.map((function(e){return t="number"==typeof(n=e).code?String(n.code):n.code,{severity:ee(n.severity),startLineNumber:n.range.start.line+1,startColumn:n.range.start.character+1,endLineNumber:n.range.end.line+1,endColumn:n.range.end.character+1,message:n.message,code:t,source:n.source};var n,t}));monaco.editor.setModelMarkers(monaco.editor.getModel(e),n,r)}))})).then(void 0,(function(e){console.error(e)}))},e}();function ee(e){switch(e){case g.Error:return monaco.MarkerSeverity.Error;case g.Warning:return monaco.MarkerSeverity.Warning;case g.Information:return monaco.MarkerSeverity.Info;case g.Hint:return monaco.MarkerSeverity.Hint;default:return monaco.MarkerSeverity.Info}}function ne(e){if(e)return{character:e.column-1,line:e.lineNumber-1}}function te(e){if(e)return new Y(e.start.line+1,e.start.character+1,e.end.line+1,e.end.character+1)}function re(e){var n=monaco.languages.CompletionItemKind;switch(e){case T.Text:return n.Text;case T.Method:return n.Method;case T.Function:return n.Function;case T.Constructor:return n.Constructor;case T.Field:return n.Field;case T.Variable:return n.Variable;case T.Class:return n.Class;case T.Interface:return n.Interface;case T.Module:return n.Module;case T.Property:return n.Property;case T.Unit:return n.Unit;case T.Value:return n.Value;case T.Enum:return n.Enum;case T.Keyword:return n.Keyword;case T.Snippet:return n.Snippet;case T.Color:return n.Color;case T.File:return n.File;case T.Reference:return n.Reference}return n.Property}function ie(e){if(e)return{range:te(e.range),text:e.newText}}var oe=function(){function e(e){this._worker=e}return Object.defineProperty(e.prototype,"triggerCharacters",{get:function(){return[".",":","<",'"',"=","/"]},enumerable:!0,configurable:!0}),e.prototype.provideCompletionItems=function(e,n,t,r){var i=e.uri;return this._worker(i).then((function(e){return e.doComplete(i.toString(),ne(n))})).then((function(t){if(t){var r=e.getWordUntilPosition(n),i=new Y(n.lineNumber,r.startColumn,n.lineNumber,r.endColumn),o=t.items.map((function(e){var n={label:e.label,insertText:e.insertText||e.label,sortText:e.sortText,filterText:e.filterText,documentation:e.documentation,detail:e.detail,range:i,kind:re(e.kind)};return e.textEdit&&(n.range=te(e.textEdit.range),n.insertText=e.textEdit.newText),e.additionalTextEdits&&(n.additionalTextEdits=e.additionalTextEdits.map(ie)),e.insertTextFormat===M.Snippet&&(n.insertTextRules=monaco.languages.CompletionItemInsertTextRule.InsertAsSnippet),n}));return{isIncomplete:t.isIncomplete,suggestions:o}}}))},e}();function ae(e){return"string"==typeof e?{value:e}:(n=e)&&"object"==typeof n&&"string"==typeof n.kind?"plaintext"===e.kind?{value:e.value.replace(/[\\`*_{}[\]()#+\-.!]/g,"\\$&")}:{value:e.value}:{value:"```"+e.language+"\n"+e.value+"\n```\n"};var n}function ue(e){if(e)return Array.isArray(e)?e.map(ae):[ae(e)]}var se=function(){function e(e){this._worker=e}return e.prototype.provideHover=function(e,n,t){var r=e.uri;return this._worker(r).then((function(e){return e.doHover(r.toString(),ne(n))})).then((function(e){if(e)return{range:te(e.range),contents:ue(e.contents)}}))},e}();function ce(e){var n=monaco.languages.DocumentHighlightKind;switch(e){case O.Read:return n.Read;case O.Write:return n.Write;case O.Text:return n.Text}return n.Text}var de=function(){function e(e){this._worker=e}return e.prototype.provideDocumentHighlights=function(e,n,t){var r=e.uri;return this._worker(r).then((function(e){return e.findDocumentHighlights(r.toString(),ne(n))})).then((function(e){if(e)return e.map((function(e){return{range:te(e.range),kind:ce(e.kind)}}))}))},e}();function fe(e){var n=monaco.languages.SymbolKind;switch(e){case W.File:return n.Array;case W.Module:return n.Module;case W.Namespace:return n.Namespace;case W.Package:return n.Package;case W.Class:return n.Class;case W.Method:return n.Method;case W.Property:return n.Property;case W.Field:return n.Field;case W.Constructor:return n.Constructor;case W.Enum:return n.Enum;case W.Interface:return n.Interface;case W.Function:return n.Function;case W.Variable:return n.Variable;case W.Constant:return n.Constant;case W.String:return n.String;case W.Number:return n.Number;case W.Boolean:return n.Boolean;case W.Array:return n.Array}return n.Function}var le=function(){function e(e){this._worker=e}return e.prototype.provideDocumentSymbols=function(e,n){var t=e.uri;return this._worker(t).then((function(e){return e.findDocumentSymbols(t.toString())})).then((function(e){if(e)return e.map((function(e){return{name:e.name,detail:"",containerName:e.containerName,kind:fe(e.kind),tags:[],range:te(e.location.range),selectionRange:te(e.location.range)}}))}))},e}(),ge=function(){function e(e){this._worker=e}return e.prototype.provideLinks=function(e,n){var t=e.uri;return this._worker(t).then((function(e){return e.findDocumentLinks(t.toString())})).then((function(e){if(e)return{links:e.map((function(e){return{range:te(e.range),url:e.target}}))}}))},e}();function he(e){return{tabSize:e.tabSize,insertSpaces:e.insertSpaces}}var pe=function(){function e(e){this._worker=e}return e.prototype.provideDocumentFormattingEdits=function(e,n,t){var r=e.uri;return this._worker(r).then((function(e){return e.format(r.toString(),null,he(n)).then((function(e){if(e&&0!==e.length)return e.map(ie)}))}))},e}(),me=function(){function e(e){this._worker=e}return e.prototype.provideDocumentRangeFormattingEdits=function(e,n,t,r){var i=e.uri;return this._worker(i).then((function(e){return e.format(i.toString(),function(e){if(e)return{start:ne(e.getStartPosition()),end:ne(e.getEndPosition())}}(n),he(t)).then((function(e){if(e&&0!==e.length)return e.map(ie)}))}))},e}(),ve=function(){function e(e){this._worker=e}return e.prototype.provideFoldingRanges=function(e,n,t){var r=e.uri;return this._worker(r).then((function(e){return e.provideFoldingRanges(r.toString(),n)})).then((function(e){if(e)return e.map((function(e){var n={start:e.startLine+1,end:e.endLine+1};return void 0!==e.kind&&(n.kind=function(e){switch(e){case d.Comment:return monaco.languages.FoldingRangeKind.Comment;case d.Imports:return monaco.languages.FoldingRangeKind.Imports;case d.Region:return monaco.languages.FoldingRangeKind.Region}}(e.kind)),n}))}))},e}();function _e(e){var n=new w(e),t=function(){for(var e=[],t=0;t<arguments.length;t++)e[t]=arguments[t];return n.getLanguageServiceWorker.apply(n,e)},r=e.languageId;monaco.languages.registerCompletionItemProvider(r,new oe(t)),monaco.languages.registerHoverProvider(r,new se(t)),monaco.languages.registerDocumentHighlightProvider(r,new de(t)),monaco.languages.registerLinkProvider(r,new ge(t)),monaco.languages.registerFoldingRangeProvider(r,new ve(t)),monaco.languages.registerDocumentSymbolProvider(r,new le(t)),"html"===r&&(monaco.languages.registerDocumentFormattingEditProvider(r,new pe(t)),monaco.languages.registerDocumentRangeFormattingEditProvider(r,new me(t)),new Z(r,t,e))}}}]);