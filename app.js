const express = require("express");
const bodyParser = require("body-parser");
const router = require("./src/routes/routes");

const app = express();
const port = process.env.PORT || 4000;

app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

app.use("/", router);

app.listen(port, "192.168.1.8", () => {
  console.log(`Server is running on port ${port}`);
});
