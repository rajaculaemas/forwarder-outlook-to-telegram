"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.AuthProvider = void 0;
const axios_1 = __importDefault(require("axios"));
const qs_1 = __importDefault(require("qs"));
const ora_1 = __importDefault(require("ora"));
require("isomorphic-fetch");
const conf_1 = __importDefault(require("conf"));
const config = new conf_1.default();
function isSuccessful(response) {
    return response.token_type !== undefined;
}
class AuthProvider {
    constructor(appData, printMessage = console.log) {
        this.app = appData;
        this.printMessage = printMessage;
        this.data = config.get('auth');
        if (this.data) {
            console.log('Loaded saved data');
        }
    }
    initDeviceCodeFlow() {
        return __awaiter(this, void 0, void 0, function* () {
            const url = `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/devicecode`;
            const { data } = yield axios_1.default.get(url, {
                params: {
                    client_id: this.app.id,
                    scope: this.app.scope
                },
                validateStatus: status => status >= 200 && status < 500
            });
            return data;
        });
    }
    pollToken(code) {
        return __awaiter(this, void 0, void 0, function* () {
            const url = `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/token`;
            const { data } = yield axios_1.default.post(url, qs_1.default.stringify({
                grant_type: 'urn:ietf:params:oauth:grant-type:device_code',
                client_id: this.app.id,
                device_code: code
            }), {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                validateStatus: status => status >= 200 && status < 500
            });
            return data;
        });
    }
    refreshTokenQuery(refreshToken) {
        return __awaiter(this, void 0, void 0, function* () {
            const url = `https://login.microsoftonline.com/${this.app.tenant}/oauth2/v2.0/token`;
            const { data } = yield axios_1.default.post(url, qs_1.default.stringify({
                grant_type: 'refresh_token',
                client_id: this.app.id,
                scope: this.app.scope,
                refresh_token: refreshToken
            }), {
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
                validateStatus: status => status >= 200 && status < 500
            });
            return data;
        });
    }
    requestToken() {
        return __awaiter(this, void 0, void 0, function* () {
            const { message, device_code, interval } = yield this.initDeviceCodeFlow();
            this.printMessage(message);
            const spinner = ora_1.default('Waiting for authorization...').start();
            let tokenData = yield this.pollToken(device_code);
            while (!isSuccessful(tokenData) && tokenData.error === 'authorization_pending') {
                // eslint-disable-next-line @typescript-eslint/no-unused-vars
                yield new Promise((resolve, reject) => {
                    setTimeout(resolve, interval * 1000);
                });
                tokenData = yield this.pollToken(device_code);
            }
            spinner.stop();
            spinner.clear();
            if (!isSuccessful(tokenData)) {
                throw new Error('Device code has expired. Please, try again.');
            }
            return {
                accessToken: tokenData.access_token,
                expireDate: Date.now() + tokenData.expires_in * 1000,
                refreshToken: tokenData.refresh_token
            };
        });
    }
    refreshToken(refreshToken) {
        return __awaiter(this, void 0, void 0, function* () {
            const tokenData = yield this.refreshTokenQuery(refreshToken);
            if (!isSuccessful(tokenData)) {
                throw new Error(tokenData.error_description);
            }
            return {
                accessToken: tokenData.access_token,
                expireDate: Date.now() + tokenData.expires_in * 1000,
                refreshToken: tokenData.refresh_token
            };
        });
    }
    /** @override */
    getAccessToken() {
        return __awaiter(this, void 0, void 0, function* () {
            if (!this.data) {
                this.data = yield this.requestToken();
                config.set('auth', this.data);
            }
            else if (this.data.expireDate <= Date.now()) {
                console.log('Token has expired');
                this.data = yield this.refreshToken(this.data.refreshToken);
                config.set('auth', this.data);
            }
            return this.data.accessToken;
        });
    }
}
exports.AuthProvider = AuthProvider;
