const path              = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = (env, argv) => {
  const isDev = argv.mode !== "production";

  return {
    entry: {
      taskpane: "./src/taskpane/taskpane.ts",
    },

    output: {
      path:     path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean:    true,
    },

    resolve: {
      extensions: [".ts", ".js"],
    },

    module: {
      rules: [
        { test: /\.ts$/,  use: "ts-loader",                    exclude: /node_modules/ },
        { test: /\.css$/, use: ["style-loader", "css-loader"],                         },
      ],
    },

    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane/taskpane.html",
        filename: "taskpane.html",
        chunks:   ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [{ from: "assets", to: "assets" }],
      }),
    ],

    devServer: {
      port:    3000,
      server:  "https",   // self-signed cert — trust it in browser once
      headers: { "Access-Control-Allow-Origin": "*" },
      hot:     true,
      open:    "/taskpane.html",   // auto-opens correct page on npm run dev
      setupMiddlewares(middlewares, devServer) {
        // Redirect / to /taskpane.html
        devServer.app.get("/", (_req, res) => res.redirect("/taskpane.html"));
        return middlewares;
      },
    },

    devtool: isDev ? "source-map" : false,
  };
};
