const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const path = require('path');

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const isServe = !!(env && env.WEBPACK_SERVE);

  return {
    devtool: 'source-map',
    entry: {
      taskpane: './src/main.jsx',
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      clean: true,
    },
    resolve: {
      extensions: ['.jsx', '.js', '.html'],
    },
    externals: {
      'office': 'Office'
    },
    module: {
      rules: [
        {
          test: /\.jsx?$/,
          exclude: /node_modules/,
          use: {
            loader: 'babel-loader',
            options: {
              presets: ['@babel/preset-env', '@babel/preset-react'],
            },
          },
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader'],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        filename: 'taskpane.html',
        template: './taskpane.html',
        chunks: ['taskpane'],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: 'manifest.xml', to: 'manifest.xml' },
        ],
      }),
    ],
    devServer: isServe
      ? {
          static: { directory: path.join(__dirname, 'dist') },
          headers: { 'Access-Control-Allow-Origin': '*' },
          server: {
            type: 'https',
            options: dev
              ? await require('office-addin-dev-certs').getHttpsServerOptions()
              : undefined,
          },
          port: 3000,
        }
      : undefined,
  };
};

