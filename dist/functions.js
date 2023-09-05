/*! For license information please see functions.js.LICENSE.txt */
!function(){function t(e){return t="function"==typeof Symbol&&"symbol"==typeof Symbol.iterator?function(t){return typeof t}:function(t){return t&&"function"==typeof Symbol&&t.constructor===Symbol&&t!==Symbol.prototype?"symbol":typeof t},t(e)}function e(){"use strict";e=function(){return r};var r={},n=Object.prototype,o=n.hasOwnProperty,a=Object.defineProperty||function(t,e,r){t[e]=r.value},i="function"==typeof Symbol?Symbol:{},c=i.iterator||"@@iterator",u=i.asyncIterator||"@@asyncIterator",s=i.toStringTag||"@@toStringTag";function f(t,e,r){return Object.defineProperty(t,e,{value:r,enumerable:!0,configurable:!0,writable:!0}),t[e]}try{f({},"")}catch(t){f=function(t,e,r){return t[e]=r}}function l(t,e,r,n){var o=e&&e.prototype instanceof d?e:d,i=Object.create(o.prototype),c=new O(n||[]);return a(i,"_invoke",{value:k(t,r,c)}),i}function h(t,e,r){try{return{type:"normal",arg:t.call(e,r)}}catch(t){return{type:"throw",arg:t}}}r.wrap=l;var p={};function d(){}function v(){}function y(){}var g={};f(g,c,(function(){return this}));var m=Object.getPrototypeOf,w=m&&m(m(R([])));w&&w!==n&&o.call(w,c)&&(g=w);var b=y.prototype=d.prototype=Object.create(g);function x(t){["next","throw","return"].forEach((function(e){f(t,e,(function(t){return this._invoke(e,t)}))}))}function E(e,r){function n(a,i,c,u){var s=h(e[a],e,i);if("throw"!==s.type){var f=s.arg,l=f.value;return l&&"object"==t(l)&&o.call(l,"__await")?r.resolve(l.__await).then((function(t){n("next",t,c,u)}),(function(t){n("throw",t,c,u)})):r.resolve(l).then((function(t){f.value=t,c(f)}),(function(t){return n("throw",t,c,u)}))}u(s.arg)}var i;a(this,"_invoke",{value:function(t,e){function o(){return new r((function(r,o){n(t,e,r,o)}))}return i=i?i.then(o,o):o()}})}function k(t,e,r){var n="suspendedStart";return function(o,a){if("executing"===n)throw new Error("Generator is already running");if("completed"===n){if("throw"===o)throw a;return{value:void 0,done:!0}}for(r.method=o,r.arg=a;;){var i=r.delegate;if(i){var c=L(i,r);if(c){if(c===p)continue;return c}}if("next"===r.method)r.sent=r._sent=r.arg;else if("throw"===r.method){if("suspendedStart"===n)throw n="completed",r.arg;r.dispatchException(r.arg)}else"return"===r.method&&r.abrupt("return",r.arg);n="executing";var u=h(t,e,r);if("normal"===u.type){if(n=r.done?"completed":"suspendedYield",u.arg===p)continue;return{value:u.arg,done:r.done}}"throw"===u.type&&(n="completed",r.method="throw",r.arg=u.arg)}}}function L(t,e){var r=e.method,n=t.iterator[r];if(void 0===n)return e.delegate=null,"throw"===r&&t.iterator.return&&(e.method="return",e.arg=void 0,L(t,e),"throw"===e.method)||"return"!==r&&(e.method="throw",e.arg=new TypeError("The iterator does not provide a '"+r+"' method")),p;var o=h(n,t.iterator,e.arg);if("throw"===o.type)return e.method="throw",e.arg=o.arg,e.delegate=null,p;var a=o.arg;return a?a.done?(e[t.resultName]=a.value,e.next=t.nextLoc,"return"!==e.method&&(e.method="next",e.arg=void 0),e.delegate=null,p):a:(e.method="throw",e.arg=new TypeError("iterator result is not an object"),e.delegate=null,p)}function S(t){var e={tryLoc:t[0]};1 in t&&(e.catchLoc=t[1]),2 in t&&(e.finallyLoc=t[2],e.afterLoc=t[3]),this.tryEntries.push(e)}function j(t){var e=t.completion||{};e.type="normal",delete e.arg,t.completion=e}function O(t){this.tryEntries=[{tryLoc:"root"}],t.forEach(S,this),this.reset(!0)}function R(t){if(t){var e=t[c];if(e)return e.call(t);if("function"==typeof t.next)return t;if(!isNaN(t.length)){var r=-1,n=function e(){for(;++r<t.length;)if(o.call(t,r))return e.value=t[r],e.done=!1,e;return e.value=void 0,e.done=!0,e};return n.next=n}}return{next:T}}function T(){return{value:void 0,done:!0}}return v.prototype=y,a(b,"constructor",{value:y,configurable:!0}),a(y,"constructor",{value:v,configurable:!0}),v.displayName=f(y,s,"GeneratorFunction"),r.isGeneratorFunction=function(t){var e="function"==typeof t&&t.constructor;return!!e&&(e===v||"GeneratorFunction"===(e.displayName||e.name))},r.mark=function(t){return Object.setPrototypeOf?Object.setPrototypeOf(t,y):(t.__proto__=y,f(t,s,"GeneratorFunction")),t.prototype=Object.create(b),t},r.awrap=function(t){return{__await:t}},x(E.prototype),f(E.prototype,u,(function(){return this})),r.AsyncIterator=E,r.async=function(t,e,n,o,a){void 0===a&&(a=Promise);var i=new E(l(t,e,n,o),a);return r.isGeneratorFunction(e)?i:i.next().then((function(t){return t.done?t.value:i.next()}))},x(b),f(b,s,"Generator"),f(b,c,(function(){return this})),f(b,"toString",(function(){return"[object Generator]"})),r.keys=function(t){var e=Object(t),r=[];for(var n in e)r.push(n);return r.reverse(),function t(){for(;r.length;){var n=r.pop();if(n in e)return t.value=n,t.done=!1,t}return t.done=!0,t}},r.values=R,O.prototype={constructor:O,reset:function(t){if(this.prev=0,this.next=0,this.sent=this._sent=void 0,this.done=!1,this.delegate=null,this.method="next",this.arg=void 0,this.tryEntries.forEach(j),!t)for(var e in this)"t"===e.charAt(0)&&o.call(this,e)&&!isNaN(+e.slice(1))&&(this[e]=void 0)},stop:function(){this.done=!0;var t=this.tryEntries[0].completion;if("throw"===t.type)throw t.arg;return this.rval},dispatchException:function(t){if(this.done)throw t;var e=this;function r(r,n){return i.type="throw",i.arg=t,e.next=r,n&&(e.method="next",e.arg=void 0),!!n}for(var n=this.tryEntries.length-1;n>=0;--n){var a=this.tryEntries[n],i=a.completion;if("root"===a.tryLoc)return r("end");if(a.tryLoc<=this.prev){var c=o.call(a,"catchLoc"),u=o.call(a,"finallyLoc");if(c&&u){if(this.prev<a.catchLoc)return r(a.catchLoc,!0);if(this.prev<a.finallyLoc)return r(a.finallyLoc)}else if(c){if(this.prev<a.catchLoc)return r(a.catchLoc,!0)}else{if(!u)throw new Error("try statement without catch or finally");if(this.prev<a.finallyLoc)return r(a.finallyLoc)}}}},abrupt:function(t,e){for(var r=this.tryEntries.length-1;r>=0;--r){var n=this.tryEntries[r];if(n.tryLoc<=this.prev&&o.call(n,"finallyLoc")&&this.prev<n.finallyLoc){var a=n;break}}a&&("break"===t||"continue"===t)&&a.tryLoc<=e&&e<=a.finallyLoc&&(a=null);var i=a?a.completion:{};return i.type=t,i.arg=e,a?(this.method="next",this.next=a.finallyLoc,p):this.complete(i)},complete:function(t,e){if("throw"===t.type)throw t.arg;return"break"===t.type||"continue"===t.type?this.next=t.arg:"return"===t.type?(this.rval=this.arg=t.arg,this.method="return",this.next="end"):"normal"===t.type&&e&&(this.next=e),p},finish:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.finallyLoc===t)return this.complete(r.completion,r.afterLoc),j(r),p}},catch:function(t){for(var e=this.tryEntries.length-1;e>=0;--e){var r=this.tryEntries[e];if(r.tryLoc===t){var n=r.completion;if("throw"===n.type){var o=n.arg;j(r)}return o}}throw new Error("illegal catch attempt")},delegateYield:function(t,e,r){return this.delegate={iterator:R(t),resultName:e,nextLoc:r},"next"===this.method&&(this.arg=void 0),p}},r}function r(t,e,r,n,o,a,i){try{var c=t[a](i),u=c.value}catch(t){return void r(t)}c.done?e(u):Promise.resolve(u).then(n,o)}function n(t){return function(){var e=this,n=arguments;return new Promise((function(o,a){var i=t.apply(e,n);function c(t){r(i,o,a,c,u,"next",t)}function u(t){r(i,o,a,c,u,"throw",t)}c(void 0)}))}}var o="https://gist.githubusercontent.com/AppFiction/750ae98fa400bb836515eee162b39057/raw/348b1216a59fef357945710cc20270b4ede089b1/market-data.json";function a(){return(a=n(e().mark((function t(){var r,n,a,i,c;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=1,t.next=4,fetch(o);case 4:return r=t.sent,t.next=7,r.json();case 7:n=t.sent,a=[],i=e().mark((function t(){var r,o;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:r=n.results[c],o=n.inputs.map((function(t){return r[t]})),a.push(o);case 3:case"end":return t.stop()}}),t)})),t.t0=e().keys(n.results);case 11:if((t.t1=t.t0()).done){t.next=16;break}return c=t.t1.value,t.delegateYield(i(),"t2",14);case 14:t.next=11;break;case 16:return Excel.run((function(t){return t.workbook.getSelectedRange().getResizedRange(a.length-1,a[0].length-1).values=a,t.sync()})).catch((function(t){console.error("Error writing data to worksheet: ",t)})),t.abrupt("return","Results Data fetched and added to the worksheet successfully!");case 20:return t.prev=20,t.t3=t.catch(1),console.error("Error fetching data: ",t.t3),t.abrupt("return","Error fetching data. Check console for details.");case 24:case"end":return t.stop()}}),t,null,[[1,20]])})))).apply(this,arguments)}function i(){return(i=n(e().mark((function t(){var r,n,a;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=3,fetch(o);case 3:return r=t.sent,t.next=6,r.json();case 6:return n=t.sent,a=[n.inputs],Excel.run((function(t){return t.workbook.getSelectedRange().getResizedRange(a.length-1,a[0].length-1).values=a,t.sync()})).catch((function(t){console.error("Error writing inputs data to worksheet: ",t)})),t.abrupt("return","Inputs data fetched and added to the worksheet successfully!");case 12:return t.prev=12,t.t0=t.catch(0),console.error("Error fetching inputs data: ",t.t0),t.abrupt("return","Error fetching inputs data. Check console for details.");case 16:case"end":return t.stop()}}),t,null,[[0,12]])})))).apply(this,arguments)}function c(){return(c=n(e().mark((function t(){var r,n;return e().wrap((function(t){for(;;)switch(t.prev=t.next){case 0:return t.prev=0,t.next=4,fetch("https://gist.githubusercontent.com/AppFiction/29f83f73b8d54cc82f693b2d4f449c24/raw/c6a8f1a69509e67593255e2071216a0a2e30df02/gistfile1.json");case 4:if((r=t.sent).ok){t.next=7;break}throw new Error("Failed to fetch data. Status: ".concat(r.status));case 7:return t.next=9,r.json();case 9:return n=t.sent,Excel.run((function(t){var e=t.workbook.getSelectedRange(),r=n.map((function(t){return[t.securityld,t.priceDate,t.baseCurrency,t.idType2,t.liqCap,t.liqFloor,t.risk,t.riskHorizon,t.riskLookbackPeriod,t.riskReturnHorizon,t.useBestPracticeRealm,t.scenarioStartDate,t.scenarioEndDate,t.columnOrder.join(", "),t.sortByColumns.join(", ")]})),o=e.getOffsetRange(1,0),a=o.getOffsetRange(1,0).getResizedRange(r.length-1,r[0].length-1),i=Object.keys(n[0]);return o.values=i,a.values=r,t.sync()})).catch((function(t){console.error("Error writing data to worksheet: ",t)})),t.abrupt("return","Data added to the worksheet successfully!");case 15:return t.prev=15,t.t0=t.catch(0),console.error("Error adding data: ",t.t0),t.abrupt("return","Error adding data. Check console for details.");case 19:case"end":return t.stop()}}),t,null,[[0,15]])})))).apply(this,arguments)}CustomFunctions.associate("FETCHMARKETDATARESULTS",(function(){return a.apply(this,arguments)})),CustomFunctions.associate("FETCHMARKETDATAINPUTS",(function(){return i.apply(this,arguments)})),CustomFunctions.associate("FETCHSUPPORTEDINPUTSBYANALYTICS",(function(){return c.apply(this,arguments)}))}();
//# sourceMappingURL=functions.js.map