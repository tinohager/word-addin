!function(){"use strict";var e,t,n,o,r={14385:function(e){e.exports=function(e,t){return t||(t={}),e?(e=String(e.__esModule?e.default:e),t.hash&&(e+=t.hash),t.maybeNeedQuotes&&/[\t\n\f\r "'=<>`]/.test(e)?'"'.concat(e,'"'):e):e}},98362:function(e,t,n){e.exports=n.p+"assets/logo-filled.png"},58394:function(e,t,n){e.exports=n.p+"1fda685b81e1123773f6.css"}},s={};function c(e){var t=s[e];if(void 0!==t)return t.exports;var n=s[e]={exports:{}};return r[e](n,n.exports,c),n.exports}c.m=r,c.n=function(e){var t=e&&e.__esModule?function(){return e.default}:function(){return e};return c.d(t,{a:t}),t},c.d=function(e,t){for(var n in t)c.o(t,n)&&!c.o(e,n)&&Object.defineProperty(e,n,{enumerable:!0,get:t[n]})},c.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}(),c.o=function(e,t){return Object.prototype.hasOwnProperty.call(e,t)},function(){var e;c.g.importScripts&&(e=c.g.location+"");var t=c.g.document;if(!e&&t&&(t.currentScript&&(e=t.currentScript.src),!e)){var n=t.getElementsByTagName("script");if(n.length)for(var o=n.length-1;o>-1&&(!e||!/^http(s?):/.test(e));)e=n[o--].src}if(!e)throw new Error("Automatic publicPath is not supported in this browser");e=e.replace(/#.*$/,"").replace(/\?.*$/,"").replace(/\/[^\/]+$/,"/"),c.p=e}(),c.b=document.baseURI||self.location.href,function(){async function e(){return await Word.run((async e=>{let t=performance.now();document.getElementById("progressbox").style.display="block",document.getElementById("progressbar").style.width="0%";const n=e.document.body.paragraphs;n.load("text"),await e.sync(),document.getElementById("progressbar").style.width="10%";const o=performance.now();console.log(`Execution time1: ${o-t} ms`),t=performance.now();const r=[];for(let e=0;e<n.items.length;e++){const t=n.items[e];if(t.text){const e=t.getRange().split([" "]);e.load("$none");try{await e.context.sync()}catch(e){console.error("error on sync1 - "+e);continue}for(let t=0;t<e.items.length;t++){const n=e.items[t].getRange().split([""]);n.load("$none"),r.push(n),n.untrack()}}}document.getElementById("progressbar").style.width="30%";const s=performance.now();console.log(`Execution time2: ${s-t} ms`),t=performance.now(),document.getElementById("progressbar").style.width="40%";try{await e.sync()}catch(e){console.error("error on sync2 - "+e)}document.getElementById("progressbar").style.width="50%";for(let e=0;e<r.length;e++){const t=r[e];try{t&&!t.isNullObject&&t.items.length>2&&(t.items[1].font.bold=!0,t.items[2].font.bold=!0)}catch(e){console.error("error on process - "+e)}}document.getElementById("progressbar").style.width="60%";const c=performance.now();console.log(`Execution time3: ${c-t} ms`),t=performance.now(),await e.sync();const l=performance.now();console.log(`Execution time4: ${l-t} ms`),document.getElementById("progressbar").style.width="100%",document.getElementById("progressbox").style.display="none"}))}async function t(){return await Word.run((async e=>{const t=performance.now();document.getElementById("progressbox").style.display="block",document.getElementById("progressbar").style.width="0%";const n=e.document.body.paragraphs;n.load("text"),await e.sync(),document.getElementById("progressbar").style.width="10%";for(let t=0;t<n.items.length;t++){const o=n.items[t],r=100*t/n.items.length;if(document.getElementById("progressbar").style.width=r+"%",o.text){const t=o.getRange().split([" "]);t.load("$none"),await t.context.sync();const n=[];for(let e=0;e<t.items.length;e++){const o=t.items[e].getRange().split([""]);o.load("$none"),n.push(o)}try{await e.sync()}catch(e){console.error("error on sync - "+e);continue}for(let e=0;e<n.length;e++){const t=n[e];t&&!t.isNullObject&&t.items.length>2&&(t.items[1].font.bold=!0,t.items[2].font.bold=!0)}}}await e.sync();const o=performance.now();console.log(`Execution time: ${o-t} ms`),document.getElementById("progressbar").style.width="100%",document.getElementById("progressbox").style.display="none"}))}Office.onReady((n=>{n.host===Office.HostType.Word&&(document.getElementById("sideload-msg").style.display="none",document.getElementById("app-body").style.display="flex",document.getElementById("run1").onclick=e,document.getElementById("run2").onclick=t)}))}(),e=c(14385),t=c.n(e),n=new URL(c(58394),c.b),o=new URL(c(98362),c.b),t()(n),t()(o)}();
//# sourceMappingURL=taskpane.js.map