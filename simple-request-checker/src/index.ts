import express, { Request } from "express";

import dotenv from "dotenv";
dotenv.config();

const app = express();

app.use(express.json());

const views = (req: Request) => {
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

const port = +process.env.PRJ_SERVER_PORT! || 3000;
app.listen(port, () => {
  console.log("start");
});
