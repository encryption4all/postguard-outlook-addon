/* eslint-disable no-undef */

const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require("webpack");
const pkg = require("./package.json");
require("dotenv").config();

const urlDev = "https://localhost:3000/";
const urlProd = process.env.ADDIN_PUBLIC_URL || "https://addin.postguard.eu/";

const requiredEnv = ["PKG_URL", "CRYPTIFY_URL", "POSTGUARD_WEBSITE_URL"];
const envDefaults = {
  PKG_URL: "https://staging.postguard.eu/pkg",
  CRYPTIFY_URL: "https://fileshare.staging.postguard.eu",
  POSTGUARD_WEBSITE_URL: "https://staging.postguard.eu",
};
const resolvedEnv = {};
for (const key of requiredEnv) {
  resolvedEnv[key] = process.env[key] || envDefaults[key];
}

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: dev ? "source-map" : false,
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: ["./src/taskpane/taskpane.ts", "./src/taskpane/taskpane.html"],
      commands: "./src/commands/commands.ts",
      launchevent: "./src/launchevent/launchevent.ts",
      "yivi-dialog": ["./src/yivi-dialog/yivi-dialog.ts", "./src/yivi-dialog/yivi-dialog.html"],
      "read-dialog": ["./src/read-dialog/read-dialog.ts", "./src/read-dialog/read-dialog.html"],
    },
    output: {
      clean: true,
    },
    experiments: {
      asyncWebAssembly: true,
      syncWebAssembly: true,
    },
    resolve: {
      extensions: [".ts", ".html", ".js", ".mjs", ".wasm"],
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: { loader: "babel-loader" },
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: {
            loader: "html-loader",
            options: {
              // yivi.min.css is copied into dist/ by CopyWebpackPlugin
              // below — leave the <link href> untouched so the browser
              // resolves it at runtime from the same directory as the
              // HTML page.
              sources: {
                urlFilter: (_attribute, value) => !/(^|\/)yivi\.min\.css$/.test(value),
              },
            },
          },
        },
        {
          test: /\.(png|jpg|jpeg|gif|ico|svg)$/,
          type: "asset/resource",
          generator: { filename: "assets/[name][ext][query]" },
        },
        {
          test: /\.wasm$/,
          type: "asset/resource",
          generator: { filename: "[name][ext]" },
        },
      ],
    },
    plugins: [
      new webpack.DefinePlugin({
        "process.env.PKG_URL": JSON.stringify(resolvedEnv.PKG_URL),
        "process.env.CRYPTIFY_URL": JSON.stringify(resolvedEnv.CRYPTIFY_URL),
        "process.env.POSTGUARD_WEBSITE_URL": JSON.stringify(resolvedEnv.POSTGUARD_WEBSITE_URL),
        // The add-in's own public origin. Needed by launchevent.ts to
        // build the Yivi dialog URL — window.location.href is unreliable
        // there because New Outlook for Mac runs the launchevent JS
        // override (JSRuntime.Url) where window.location is an Office-
        // internal URL, not the add-in origin.
        "process.env.ADDIN_PUBLIC_URL": JSON.stringify(dev ? urlDev : urlProd),
        // Stamped into the X-PostGuard-Client-Version header so PKG /
        // Cryptify dashboards can attribute metrics per release. Sourced
        // from package.json, which release-please bumps on every release.
        "process.env.ADDIN_VERSION": JSON.stringify(pkg.version),
      }),
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "launchevent.html",
        template: "./src/launchevent/launchevent.html",
        chunks: ["launchevent"],
      }),
      new HtmlWebpackPlugin({
        filename: "yivi-dialog.html",
        template: "./src/yivi-dialog/yivi-dialog.html",
        chunks: ["polyfill", "yivi-dialog"],
      }),
      new HtmlWebpackPlugin({
        filename: "read-dialog.html",
        template: "./src/read-dialog/read-dialog.html",
        chunks: ["polyfill", "read-dialog"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets/*", to: "assets/[name][ext][query]" },
          {
            from: "node_modules/@privacybydesign/yivi-css/dist/yivi.min.css",
            to: "yivi.min.css",
          },
          {
            from: "manifest*.xml",
            to: "[name][ext]",
            transform(content) {
              if (dev) return content;
              return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
            },
          },
        ],
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options:
          env.WEBPACK_BUILD || options.https !== undefined
            ? options.https
            : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};
