"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
const express_1 = __importDefault(require("express"));
const dotenv_1 = __importDefault(require("dotenv"));
dotenv_1.default.config();
const app = (0, express_1.default)();
app.use(express_1.default.json());
const views = (req) => {
    const { httpVersion, method, originalUrl, headers, query, path, body } = req;
    return { httpVersion, method, originalUrl, headers, query, path, body };
};
app.get("/*", (req, _a, _b) => {
    console.log(views(req));
});
app.post("/*", (req, _a, _b) => {
    console.log(views(req));
});
app.delete("/*", (req, _a, _b) => {
    console.log(views(req));
});
app.put("/*", (req, _a, _b) => {
    console.log(views(req));
});
app.patch("/*", (req, _a, _b) => {
    console.log(views(req));
});
const port = +process.env.PRJ_SERVER_PORT || 3000;
app.listen(port, () => {
    console.log("start");
});
