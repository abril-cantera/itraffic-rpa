const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const path = require('path');
const fs = require('fs');

const urlDev = 'https://localhost:3000/';
const urlProd = 'https://happy-flower-09b6bd81e.4.azurestaticapps.net/';

// Generate self-signed certificate for local development if not exists
function getDevServerHttpsConfig() {
  // First check Office add-in dev certs location
  const officeCertPath = require('os').homedir() + '/.office-addin-dev-certs';
  const officeKeyFile = path.join(officeCertPath, 'localhost.key');
  const officeCertFile = path.join(officeCertPath, 'localhost.crt');
  
  if (fs.existsSync(officeKeyFile) && fs.existsSync(officeCertFile)) {
    console.log('📜 Using Office Add-in dev certificates');
    return {
      key: fs.readFileSync(officeKeyFile),
      cert: fs.readFileSync(officeCertFile),
    };
  }
  
  // Then check local certs folder
  const certPath = path.join(__dirname, 'certs');
  const keyFile = path.join(certPath, 'localhost-key.pem');
  const certFile = path.join(certPath, 'localhost.pem');
  
  if (fs.existsSync(keyFile) && fs.existsSync(certFile)) {
    console.log('📜 Using local certificates');
    return {
      key: fs.readFileSync(keyFile),
      cert: fs.readFileSync(certFile),
    };
  }
  
  // Use default self-signed cert from webpack
  console.log('📜 Using webpack default certificates');
  return true;
}

module.exports = async (env, options) => {
  const dev = options.mode === 'development';
  const buildType = dev ? 'dev' : 'prod';
  const url = dev ? urlDev : urlProd;

  return {
    devtool: dev ? 'eval-source-map' : 'source-map',
    entry: {
      taskpane: './src/taskpane/taskpane.ts',
      commands: './src/commands/commands.ts',
    },
    output: {
      path: path.resolve(__dirname, 'dist'),
      filename: '[name].js',
      clean: true,
    },
    resolve: {
      extensions: ['.ts', '.tsx', '.html', '.js'],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: 'ts-loader',
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ['style-loader', 'css-loader'],
        },
        {
          test: /\.(png|jpg|jpeg|gif|svg)$/,
          type: 'asset/resource',
          generator: {
            filename: 'assets/[name][ext]',
          },
        },
      ],
    },
    plugins: [
      // Generate index.html for Azure Static Web Apps requirement
      new HtmlWebpackPlugin({
        filename: 'index.html',
        templateContent: `
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>iTraffic RPA - Outlook Add-in</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            display: flex;
            justify-content: center;
            align-items: center;
            height: 100vh;
            margin: 0;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        .container {
            text-align: center;
            padding: 2rem;
            background: rgba(255, 255, 255, 0.1);
            border-radius: 10px;
            backdrop-filter: blur(10px);
        }
        h1 { margin-bottom: 1rem; }
        p { margin-bottom: 1.5rem; }
        a {
            color: white;
            text-decoration: none;
            padding: 10px 20px;
            background: rgba(255, 255, 255, 0.2);
            border-radius: 5px;
            transition: background 0.3s;
        }
        a:hover {
            background: rgba(255, 255, 255, 0.3);
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>📧 iTraffic RPA - Outlook Add-in</h1>
        <p>Este es un complemento de Outlook para crear reservas en iTraffic automáticamente.</p>
        <p>Para usar este add-in, instálalo desde Outlook.</p>
        <a href="manifest.xml">Ver Manifest</a>
    </div>
</body>
</html>
        `,
        inject: false,
      }),
      new HtmlWebpackPlugin({
        template: './src/taskpane/taskpane.html',
        filename: 'taskpane.html',
        chunks: ['taskpane'],
      }),
      new HtmlWebpackPlugin({
        template: './src/commands/commands.html',
        filename: 'commands.html',
        chunks: ['commands'],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: 'assets',
            to: 'assets',
          },
          {
            from: 'src/login/login.html',
            to: 'login.html',
          },
          {
            from: 'src/login/login.js',
            to: 'login.js',
          },
          {
            from: 'manifest.xml',
            to: 'manifest.xml',
            transform: (content) => {
              if (dev) {
                return content;
              }
              // Replace localhost URLs with production URLs
              return content
                .toString()
                .replace(/https:\/\/localhost:3000/g, url.replace(/\/$/, ''));
            },
          },
        ],
      }),
    ],
    devServer: {
      static: {
        directory: path.join(__dirname, 'dist'),
      },
      headers: {
        'Access-Control-Allow-Origin': '*',
      },
      server: {
        type: 'https',
        options: getDevServerHttpsConfig(),
      },
      port: 3000,
      hot: true,
      // Allow connections from Outlook
      allowedHosts: 'all',
    },
  };
};
