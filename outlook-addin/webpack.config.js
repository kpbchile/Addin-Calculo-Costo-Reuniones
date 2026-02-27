const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    launchevent: "./src/launchevent/launchevent.js",
    commands: "./src/commands/commands.js",
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js",
    clean: true,
  },
  devServer: {
    static: {
      directory: path.join(__dirname, "dist"),
    },
    headers: {
      "Access-Control-Allow-Origin": "*",
    },
    server: "https",
    port: 3000,
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "index.html",
      template: "./src/index.html",
      chunks: [],
    }),
    new HtmlWebpackPlugin({
      filename: "commands.html",
      template: "./src/commands/commands.html",
      chunks: ["commands"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: "manifest.xml", to: "manifest.xml" },
      ],
    }),
  ],
};
