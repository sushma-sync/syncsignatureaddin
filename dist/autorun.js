!function(){var e={14385:function(e){"use strict";e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},5373:function(e,t,n){"use strict";e.exports=n.p+"de212c86fba242a584d5.js"},25464:function(e,t,n){"use strict";e.exports=n.p+"c06ba5e7b70910052e1c.js"}},t={};function n(o){var i=t[o];if(void 0!==i)return i.exports;var r=t[o]={exports:{}};return e[o](r,r.exports,n),r.exports}n.m=e,n.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return n.d(t,{a:t}),t},n.d=function(e,t){for(var o in t)n.o(t,o)&&!n.o(e,o)&&Object.defineProperty(e,o,{enumerable:!0,get:t[o]})},n.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),n.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;n.g.importScripts&&(e=n.g.location+"");var t=n.g.document;if(!e&&t&&(t.currentScript&&"SCRIPT"===t.currentScript.tagName.toUpperCase()&&(e=t.currentScript.src),!e)){var o=t.getElementsByTagName("script");if(o.length)for(var i=o.length-1;i>-1&&(!e||!/^http(s?):/.test(e));)e=o[i--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),n.p=e}(),n.b=document.baseURI||self.location.href,function(){function e(e,n,o){var i=function(e){return"reply"===e?Office.context.roamingSettings.get("reply"):"forward"===e?Office.context.roamingSettings.get("forward"):Office.context.roamingSettings.get("newMail")}(e),r=function(e,n){return"templateB"===e?function(e){var n="";return t(e.greeting)&&(n+=e.greeting+"<br/>"),n+="<table>",n+="<tr>",n+="<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='https://officedev.github.io/Office-Add-in-samples/Samples/outlook-set-signature/assets/sample-logo.png' alt='Logo' /></td>",n+="<td style='padding-left: 5px;'>",n+="<strong>"+e.name+"</strong>",n+=t(e.pronoun)?"&nbsp;"+e.pronoun:"",n+="<br/>",n+=e.email+"<br/>",n+=t(e.phone)?e.phone+"<br/>":"",n+="</td>",n+="</tr>",{signature:n+="</table>",logoBase64:null,logoFileName:null}}(n):"templateC"===e?function(e){var n="";return t(e.greeting)&&(n+=e.greeting+"<br/>"),{signature:n+=e.name,logoBase64:null,logoFileName:null}}(n):function(e){var n="sample-logo.png",o="";return t(e.greeting)&&(o+=e.greeting+"<br/>"),o+="<table>",o+="<tr>",o+="<td style='border-right: 1px solid #000000; padding-right: 5px;'><img src='cid:"+n+"' alt='MS Logo' width='24' height='24' /></td>",o+="<td style='padding-left: 5px;'>",o+="<strong>"+e.name+"</strong>",o+=t(e.pronoun)?"&nbsp;"+e.pronoun:"",o+="<br/>",o+=t(e.job)?e.job+"<br/>":"",o+=e.email+"<br/>",o+=t(e.phone)?e.phone+"<br/>":"",o+="</td>",o+="</tr>",{signature:o+="</table>",logoBase64:"iVBORw0KGgoAAAANSUhEUgAAACIAAAAiCAYAAAA6RwvCAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsQAAA7EAZUrDhsAAAEeSURBVFhHzdhBEoIwDIVh4EoeQJd6YrceQM+kvo5hQNokLymO/4aF0/ajlBl1fL4bEp0uj3K9XQ/lGi0MEcB3UdD0uVK1EEj7TIuGeBaKYCgIswCLcUMid8mMcUEiCMk71oRYE+Etsd4UD0aFeBBSFtOEMAgpg6lCIggpitlAMggpgllBeiAkFjNDeiIkBlMgeyAkL6Z6WJdlEJJnjvF4vje/BvRALNN23tyRXzVpd22dHSZtLhjMHemB8cxRINZZyGCssbL2vCN7YLwItHo0PTEMAm3OSA8Mi0DVw5rBRBCoCkERTBSBmhDEYDII5PqlZy1iZSGQuiOSZ6JW3rEuCIpgmDFuCGImZuEUBHkWiOweDUHaQhEE+pM/aobhBZaOpYLJeeeoAAAAAElFTkSuQmCC",logoFileName:n}}(n)}(i,n);!function(e,n,o){!0===t(e.logoBase64)?Office.context.mailbox.item.addFileAttachmentFromBase64Async(e.logoBase64,e.logoFileName,{isInline:!0},(function(t){Office.context.mailbox.item.body.setSignatureAsync(e.signature,{coercionType:"html",asyncContext:n},(function(e){e.asyncContext.completed()}))})):Office.context.mailbox.item.body.setSignatureAsync(e.signature,{coercionType:"html",asyncContext:n},(function(e){e.asyncContext.completed()}))}(r,o)}function t(e){return null!=e&&""!==e}Office.actions.associate("checkSignature",(function(t){var n=Office.context.roamingSettings.get("user_info");if(n){var o=JSON.parse(n);Office.context.mailbox.item.getComposeTypeAsync?Office.context.mailbox.item.getComposeTypeAsync({asyncContext:{user_info:o,eventObj:t}},(function(t){"succeeded"===t.status&&e(t.value.composeType,t.asyncContext.user_info,t.asyncContext.eventObj)})):e("newMail",JSON.parse(n),t)}else Office.context.mailbox.item.notificationMessages.addAsync("fd90eb33431b46f58a68720c36154b4a",{type:"insightMessage",message:"Please set your signature with the Office Add-ins sample.",icon:"Icon.16x16",actions:[{actionType:"showTaskPane",actionText:"Set signatures",commandId:"appointment"==Office.context.mailbox.item.itemType?"MRCS_TpBtn1":"MRCS_TpBtn0",contextData:"{''}"}]})}))}(),function(){"use strict";var e=n(14385),t=n.n(e),o=new URL(n(5373),n.b),i=new URL(n(25464),n.b);t()(o),t()(i)}()}();
//# sourceMappingURL=autorun.js.map