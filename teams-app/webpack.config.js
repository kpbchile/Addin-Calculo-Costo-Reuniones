const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
  entry: {
    tab: "./src/tab.js",
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
    port: 3001,
  },
  plugins: [
    new HtmlWebpackPlugin({
      filename: "tab.html",
      template: "./src/tab.html",
      chunks: ["tab"],
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: "assets", to: "assets" },
        { from: "manifest.json", to: "manifest.json" },
      ],
    }),
  ],
};
