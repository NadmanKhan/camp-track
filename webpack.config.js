// Docs: https://webpack.js.org/configuration/
/**
 * @type {import('path')}
 */
const path = require('path');

/**
 * @type {import('webpack').WebpackPluginInstance}
 */
const GasPlugin = require('gas-webpack-plugin');

/**
 * @type {import('webpack').Configuration}
 */
module.exports = {
  mode: 'development',
  entry: {
    main: path.resolve(__dirname, 'src', 'main.ts'),
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'bundle.js',
  },  resolve: {
    extensions: [".ts", ".js"],
  },
  module: {
    rules: [
      {
        test: /\.ts$/,
        use: "ts-loader",
      },
    ],
  },
  devtool: 'source-map',
  optimization: {
    minimize: false, // disable uglify to keep code readable for GAS
  },
  plugins: [
    new GasPlugin(),
  ],
};