/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/middle-tier/msgraph-helper.ts":
/*!*******************************************!*\
  !*** ./src/middle-tier/msgraph-helper.ts ***!
  \*******************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getGraphData: () => (/* binding */ getGraphData),
/* harmony export */   getUserData: () => (/* binding */ getUserData)
/* harmony export */ });
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! https */ "https");
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(https__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var _ssoauth_helper__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! ./ssoauth-helper */ "./src/middle-tier/ssoauth-helper.ts");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! http-errors */ "http-errors");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(http_errors__WEBPACK_IMPORTED_MODULE_2__);
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
*/




/* global process */

const domain = "graph.microsoft.com";
const version = "v1.0";
async function getUserData(req, res, next) {
  const authorization = req.get("Authorization");
  await (0,_ssoauth_helper__WEBPACK_IMPORTED_MODULE_1__.getAccessToken)(authorization).then(async graphTokenResponse => {
    if (graphTokenResponse && (graphTokenResponse.claims || graphTokenResponse.error)) {
      res.send(graphTokenResponse);
    } else {
      const graphToken = graphTokenResponse.access_token;
      const graphUrlSegment = process.env.GRAPH_URL_SEGMENT || "/me";
      const graphQueryParamSegment = process.env.QUERY_PARAM_SEGMENT || "";
      const graphData = await getGraphData(graphToken, graphUrlSegment, graphQueryParamSegment);

      // If Microsoft Graph returns an error, such as invalid or expired token,
      // there will be a code property in the returned object set to a HTTP status (e.g. 401).
      // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
      if (graphData.code) {
        next(http_errors__WEBPACK_IMPORTED_MODULE_2__(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
      } else {
        res.send(graphData);
      }
    }
  }).catch(err => {
    res.status(401).send(err.message);
    return;
  });
}
async function getGraphData(accessToken, apiUrl, queryParams) {
  return new Promise((resolve, reject) => {
    const options = {
      host: domain,
      path: "/" + version + apiUrl + queryParams,
      method: "GET",
      headers: {
        "Content-Type": "application/json",
        Accept: "application/json",
        Authorization: "Bearer " + accessToken,
        "Cache-Control": "private, no-cache, no-store, must-revalidate",
        Expires: "-1",
        Pragma: "no-cache"
      }
    };
    https__WEBPACK_IMPORTED_MODULE_0__.get(options, response => {
      let body = "";
      response.on("data", d => {
        body += d;
      });
      response.on("end", () => {
        // The response from the OData endpoint might be an error, say a
        // 401 if the endpoint requires an access token and it was invalid
        // or expired. But a message is not an error in the call of https.get,
        // so the "on('error', reject)" line below isn't triggered.
        // So, the code distinguishes success (200) messages from error
        // messages and sends a JSON object to the caller with either the
        // requested OData or error information.

        let error;
        if (response.statusCode === 200) {
          let parsedBody = JSON.parse(body);
          resolve(parsedBody);
        } else {
          error = new Error();
          error.code = response.statusCode;
          error.message = response.statusMessage;

          // The error body sometimes includes an empty space
          // before the first character, remove it or it causes an error.
          body = body.trim();
          error.bodyCode = JSON.parse(body).error.code;
          error.bodyMessage = JSON.parse(body).error.message;
          resolve(error);
        }
      });
    }).on("error", reject);
  });
}

/***/ }),

/***/ "./src/middle-tier/ssoauth-helper.ts":
/*!*******************************************!*\
  !*** ./src/middle-tier/ssoauth-helper.ts ***!
  \*******************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   getAccessToken: () => (/* binding */ getAccessToken),
/* harmony export */   validateJwt: () => (/* binding */ validateJwt)
/* harmony export */ });
/* harmony import */ var node_fetch__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! node-fetch */ "node-fetch");
/* harmony import */ var node_fetch__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(node_fetch__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var form_urlencoded__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! form-urlencoded */ "form-urlencoded");
/* harmony import */ var form_urlencoded__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(form_urlencoded__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var jsonwebtoken__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! jsonwebtoken */ "jsonwebtoken");
/* harmony import */ var jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(jsonwebtoken__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var jwks_rsa__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! jwks-rsa */ "jwks-rsa");
/* harmony import */ var jwks_rsa__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(jwks_rsa__WEBPACK_IMPORTED_MODULE_3__);
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */






/* global process, console */

const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";
async function getAccessToken(authorization) {
  if (!authorization) {
    let error = new Error("No Authorization header was found.");
    return Promise.reject(error);
  } else {
    const scopeName = process.env.SCOPE || "User.Read";
    const [, /* schema */assertion] = authorization.split(" ");
    const tokenScopes = jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default().decode(assertion).scp.split(" ");
    const accessAsUserScope = tokenScopes.find(scope => scope === "access_as_user");
    if (!accessAsUserScope) {
      throw new Error("Missing access_as_user");
    }
    const formParams = {
      client_id: process.env.CLIENT_ID,
      client_secret: process.env.CLIENT_SECRET,
      grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
      assertion: assertion,
      requested_token_use: "on_behalf_of",
      scope: [scopeName].join(" ")
    };
    const stsDomain = "https://login.microsoftonline.com";
    const tenant = "common";
    const tokenURLSegment = "oauth2/v2.0/token";
    const encodedForm = form_urlencoded__WEBPACK_IMPORTED_MODULE_1___default()(formParams);
    const tokenResponse = await node_fetch__WEBPACK_IMPORTED_MODULE_0___default()(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
      method: "POST",
      body: encodedForm,
      headers: {
        Accept: "application/json",
        "Content-Type": "application/x-www-form-urlencoded"
      }
    });
    const json = await tokenResponse.json();
    return json;
  }
}
function validateJwt(req, res, next) {
  const authHeader = req.headers.authorization;
  if (authHeader) {
    const token = authHeader.split(" ")[1];
    const validationOptions = {
      audience: process.env.CLIENT_ID
    };
    jsonwebtoken__WEBPACK_IMPORTED_MODULE_2___default().verify(token, getSigningKeys, validationOptions, err => {
      if (err) {
        console.log(err);
        return res.sendStatus(403);
      }
      next();
    });
  }
}
function getSigningKeys(header, callback) {
  var client = new jwks_rsa__WEBPACK_IMPORTED_MODULE_3__.JwksClient({
    jwksUri: DISCOVERY_KEYS_ENDPOINT
  });
  client.getSigningKey(header.kid, function (err, key) {
    callback(null, key.getPublicKey());
  });
}

/***/ }),

/***/ "cookie-parser":
/*!********************************!*\
  !*** external "cookie-parser" ***!
  \********************************/
/***/ ((module) => {

module.exports = require("cookie-parser");

/***/ }),

/***/ "dotenv":
/*!*************************!*\
  !*** external "dotenv" ***!
  \*************************/
/***/ ((module) => {

module.exports = require("dotenv");

/***/ }),

/***/ "express":
/*!**************************!*\
  !*** external "express" ***!
  \**************************/
/***/ ((module) => {

module.exports = require("express");

/***/ }),

/***/ "form-urlencoded":
/*!**********************************!*\
  !*** external "form-urlencoded" ***!
  \**********************************/
/***/ ((module) => {

module.exports = require("form-urlencoded");

/***/ }),

/***/ "http-errors":
/*!******************************!*\
  !*** external "http-errors" ***!
  \******************************/
/***/ ((module) => {

module.exports = require("http-errors");

/***/ }),

/***/ "jsonwebtoken":
/*!*******************************!*\
  !*** external "jsonwebtoken" ***!
  \*******************************/
/***/ ((module) => {

module.exports = require("jsonwebtoken");

/***/ }),

/***/ "jwks-rsa":
/*!***************************!*\
  !*** external "jwks-rsa" ***!
  \***************************/
/***/ ((module) => {

module.exports = require("jwks-rsa");

/***/ }),

/***/ "morgan":
/*!*************************!*\
  !*** external "morgan" ***!
  \*************************/
/***/ ((module) => {

module.exports = require("morgan");

/***/ }),

/***/ "node-fetch":
/*!*****************************!*\
  !*** external "node-fetch" ***!
  \*****************************/
/***/ ((module) => {

module.exports = require("node-fetch");

/***/ }),

/***/ "office-addin-dev-certs":
/*!*****************************************!*\
  !*** external "office-addin-dev-certs" ***!
  \*****************************************/
/***/ ((module) => {

module.exports = require("office-addin-dev-certs");

/***/ }),

/***/ "path":
/*!***********************!*\
  !*** external "path" ***!
  \***********************/
/***/ ((module) => {

module.exports = require("path");

/***/ }),

/***/ "https":
/*!************************!*\
  !*** external "https" ***!
  \************************/
/***/ ((module) => {

module.exports = require("https");

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/compat get default export */
/******/ 	(() => {
/******/ 		// getDefaultExport function for compatibility with non-harmony modules
/******/ 		__webpack_require__.n = (module) => {
/******/ 			var getter = module && module.__esModule ?
/******/ 				() => (module['default']) :
/******/ 				() => (module);
/******/ 			__webpack_require__.d(getter, { a: getter });
/******/ 			return getter;
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry need to be wrapped in an IIFE because it need to be isolated against other modules in the chunk.
(() => {
/*!********************************!*\
  !*** ./src/middle-tier/app.ts ***!
  \********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_0__ = __webpack_require__(/*! http-errors */ "http-errors");
/* harmony import */ var http_errors__WEBPACK_IMPORTED_MODULE_0___default = /*#__PURE__*/__webpack_require__.n(http_errors__WEBPACK_IMPORTED_MODULE_0__);
/* harmony import */ var path__WEBPACK_IMPORTED_MODULE_1__ = __webpack_require__(/*! path */ "path");
/* harmony import */ var path__WEBPACK_IMPORTED_MODULE_1___default = /*#__PURE__*/__webpack_require__.n(path__WEBPACK_IMPORTED_MODULE_1__);
/* harmony import */ var cookie_parser__WEBPACK_IMPORTED_MODULE_2__ = __webpack_require__(/*! cookie-parser */ "cookie-parser");
/* harmony import */ var cookie_parser__WEBPACK_IMPORTED_MODULE_2___default = /*#__PURE__*/__webpack_require__.n(cookie_parser__WEBPACK_IMPORTED_MODULE_2__);
/* harmony import */ var morgan__WEBPACK_IMPORTED_MODULE_3__ = __webpack_require__(/*! morgan */ "morgan");
/* harmony import */ var morgan__WEBPACK_IMPORTED_MODULE_3___default = /*#__PURE__*/__webpack_require__.n(morgan__WEBPACK_IMPORTED_MODULE_3__);
/* harmony import */ var express__WEBPACK_IMPORTED_MODULE_4__ = __webpack_require__(/*! express */ "express");
/* harmony import */ var express__WEBPACK_IMPORTED_MODULE_4___default = /*#__PURE__*/__webpack_require__.n(express__WEBPACK_IMPORTED_MODULE_4__);
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_5__ = __webpack_require__(/*! https */ "https");
/* harmony import */ var https__WEBPACK_IMPORTED_MODULE_5___default = /*#__PURE__*/__webpack_require__.n(https__WEBPACK_IMPORTED_MODULE_5__);
/* harmony import */ var office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_6__ = __webpack_require__(/*! office-addin-dev-certs */ "office-addin-dev-certs");
/* harmony import */ var office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_6___default = /*#__PURE__*/__webpack_require__.n(office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_6__);
/* harmony import */ var _msgraph_helper__WEBPACK_IMPORTED_MODULE_7__ = __webpack_require__(/*! ./msgraph-helper */ "./src/middle-tier/msgraph-helper.ts");
/* harmony import */ var _ssoauth_helper__WEBPACK_IMPORTED_MODULE_8__ = __webpack_require__(/*! ./ssoauth-helper */ "./src/middle-tier/ssoauth-helper.ts");
/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */

if (true) {
  (__webpack_require__(/*! dotenv */ "dotenv").config)();
}










/* global console, process, require, __dirname */

const app = express__WEBPACK_IMPORTED_MODULE_4___default()();
const port = process.env.API_PORT || "3000";
app.set("port", port);

// view engine setup
app.set("views", path__WEBPACK_IMPORTED_MODULE_1__.join(__dirname, "views"));
app.set("view engine", "pug");
app.use(morgan__WEBPACK_IMPORTED_MODULE_3__("dev"));
app.use(express__WEBPACK_IMPORTED_MODULE_4___default().json());
app.use(express__WEBPACK_IMPORTED_MODULE_4___default().urlencoded({
  extended: false
}));
app.use(cookie_parser__WEBPACK_IMPORTED_MODULE_2__());

/* Turn off caching when developing */
if (true) {
  app.use(express__WEBPACK_IMPORTED_MODULE_4___default()["static"](path__WEBPACK_IMPORTED_MODULE_1__.join(process.cwd(), "dist"), {
    etag: false
  }));
  app.use(function (req, res, next) {
    res.header("Cache-Control", "private, no-cache, no-store, must-revalidate");
    res.header("Expires", "-1");
    res.header("Pragma", "no-cache");
    next();
  });
} else {}
const indexRouter = express__WEBPACK_IMPORTED_MODULE_4___default().Router();
indexRouter.get("/", function (req, res) {
  res.render("/taskpane.html");
});
app.use("/", indexRouter);

// Middle-tier API calls
// listen for 'ping' to verify service is running
// Un comment for development debugging, but un needed for production deployment
// app.get("/ping", function (req: any, res: any) {
//   res.send(process.platform);
// });

//app.get("/getuserdata", validateJwt, getUserData);
app.get("/getuserdata", _ssoauth_helper__WEBPACK_IMPORTED_MODULE_8__.validateJwt, _msgraph_helper__WEBPACK_IMPORTED_MODULE_7__.getUserData);

// Get the client side task pane files requested
app.get("/taskpane.html", async (req, res) => {
  return res.sendfile("taskpane.html");
});
app.get("/fallbackauthdialog.html", async (req, res) => {
  return res.sendfile("fallbackauthdialog.html");
});

// Catch 404 and forward to error handler
app.use(function (req, res, next) {
  next(http_errors__WEBPACK_IMPORTED_MODULE_0__(404));
});

// error handler
app.use(function (err, req, res) {
  // set locals, only providing error in development
  res.locals.message = err.message;
  res.locals.error = req.app.get("env") === "development" ? err : {};

  // render the error page
  res.status(err.status || 500);
  res.render("error");
});
(0,office_addin_dev_certs__WEBPACK_IMPORTED_MODULE_6__.getHttpsServerOptions)().then(options => {
  https__WEBPACK_IMPORTED_MODULE_5___default().createServer(options, app).listen(port, () => console.log(`Server running on ${port} in ${"development"} mode`));
});
})();

/******/ })()
;
//# sourceMappingURL=middletier.js.map