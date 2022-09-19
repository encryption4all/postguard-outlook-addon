/* eslint-disable no-undef */

const devCerts = require('office-addin-dev-certs')
const { CleanWebpackPlugin } = require('clean-webpack-plugin')
const CopyWebpackPlugin = require('copy-webpack-plugin')
const HtmlWebpackPlugin = require('html-webpack-plugin')
const path = require('path')
//const ReplaceInFileWebpackPlugin = require('replace-in-file-webpack-plugin')
const preprocess = require('svelte-preprocess')

const webpack = require('webpack')

const urlDev = 'localhost:3000/'
const urlProd = 'irmaseal.z6.web.core.windows.net/' // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION
//const appIdProd = '6ee2a054-1d61-405d-8e5d-c2daf25c5833' // CHANGE TO APP ID used in App registration of RU account

module.exports = async (env, options) => {
  const dev = options.mode === 'development'

  const buildType = dev ? 'dev' : 'prod'
  const config = {
    devtool: 'source-map',
    entry: {
      polyfill: ['core-js/stable', 'regenerator-runtime/runtime'],
      utils: './src/helpers/utils.ts',
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts',
      decrypt: './src/decryptdialog/decrypt.ts',
      fallbackauthdialog: './src/helpers/fallbackauthdialog.ts',
      attributes: './src/dialogs/attributes.ts',
      settings: './src/taskpane/settings.ts'
    },
    experiments: { syncWebAssembly: true, topLevelAwait: true },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js', '.mjs', '.svelte'],
      alias: {
        process: 'process/browser',
        stream: 'stream-browserify',
        zlib: 'browserify-zlib',
        crypto: 'crypto-browserify',
        svelte: path.resolve('node_modules', 'svelte')
      },
      fallback: { https: false, http: false },
      mainFields: ['svelte', 'browser', 'module', 'main']
    },
    module: {
      rules: [
        {
          test: /\.ts$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: ['@babel/preset-typescript'],
              plugins: ['@babel/plugin-syntax-top-level-await']
            }
          }
        },
        {
          test: /\.tsx?$/,
          exclude: /node_modules/,
          use: 'ts-loader'
        },
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: 'html-loader'
        },
        {
          test: /\.(png|jpg|jpeg|gif)$/,
          loader: 'file-loader',
          options: {
            name: '[path][name].[ext]'
          }
        },
        {
          test: /\.(svelte)$/,
          use: {
            loader: 'svelte-loader',
            options: { preprocess: preprocess({ postcss: true }) }
          }
        },
        {
          test: /node_modules\/svelte\/.*\.mjs$/,
          resolve: {
            fullySpecified: false
          }
        },
        {
          test: /\.(woff(2)?|ttf|eot)(\?v=\d+\.\d+\.\d+)?$/,
          use: [
            {
              loader: 'file-loader',
              options: {
                name: '[name].[ext]',
                outputPath: 'fonts/'
              }
            }
          ]
        },
        { test: /\.(css)$/, use: ['style-loader', 'css-loader'] }
      ]
    },
    plugins: [
      new CopyWebpackPlugin({ patterns: [{ from: 'assets', to: 'assets' }] }),
      new CopyWebpackPlugin({ patterns: [{ from: 'fonts', to: 'fonts' }] }),
      new CopyWebpackPlugin({ patterns: [{ from: 'locales', to: 'locales' }] }),
      new webpack.ProvidePlugin({
        process: 'process/browser',
        Buffer: ['buffer', 'Buffer']
      }),
      new CleanWebpackPlugin(),
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './src/taskpane/taskpane.html',
        chunks: ['polyfill', 'taskpane']
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            to: 'taskpane.css',
            from: './src/taskpane/taskpane.css'
          },
          {
            to: 'decrypt.css',
            from: './src/decryptdialog/decrypt.css'
          },
          {
            to: '[name].' + buildType + '.[ext]',
            from: 'manifest*.xml',
            transform(content) {
              if (dev) {
                return content
              } else {
                return content
                  .toString()
                  .replace(new RegExp(urlDev, 'g'), urlProd)
              }
            }
          }
        ]
      }),
      new HtmlWebpackPlugin({
        filename: 'commands.html',
        template: './src/commands/commands.html',
        chunks: ['polyfill', 'commands']
      }),
      new HtmlWebpackPlugin({
        filename: 'decrypt.html',
        template: './src/decryptdialog/decrypt.html',
        chunks: ['polyfill', 'decrypt']
      }),
      new HtmlWebpackPlugin({
        filename: 'success.html',
        template: './src/dialogs/success.html',
        chunks: ['polyfill', 'decrypt']
      }),
      new HtmlWebpackPlugin({
        filename: 'bcc.html',
        template: './src/dialogs/bcc.html',
        chunks: ['polyfill', 'commands']
      }),
      new HtmlWebpackPlugin({
        filename: 'attributes.html',
        template: './src/dialogs/attributes.html',
        chunks: ['polyfill', 'attributes']
      }),
      new HtmlWebpackPlugin({
        filename: 'fallbackauthdialog.html',
        template: './src/helpers/fallbackauthdialog.html',
        chunks: ['polyfill', 'fallbackauthdialog']
      }),
      new HtmlWebpackPlugin({
        filename: 'attributes.html',
        template: './src/dialogs/attributes.html',
        chunks: ['polyfill', 'attributes']
      }),
      new HtmlWebpackPlugin({
        filename: 'settings.html',
        template: './src/taskpane/settings.html',
        chunks: ['polyfill', 'settings']
      })
    ],
    devServer: {
      headers: {
        'Access-Control-Allow-Origin': '*'
      },
      https:
        options.https !== undefined
          ? options.https
          : await devCerts.getHttpsServerOptions(),
      port: process.env.npm_package_config_dev_server_port || 3000,
      disableHostCheck: true
    },
    output: {
      publicPath: ''
    }
  }

  return config
}
