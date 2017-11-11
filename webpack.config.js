var path = require("path");
var webpack = require("webpack");
const UglifyJSPlugin = require('uglifyjs-webpack-plugin');
var CopyWebpackPlugin = require('copy-webpack-plugin');

module.exports = {
    target: "web",
    entry: {
        App: "./src/scripts/App.tsx",
        SettingsPanel: "./src/scripts/SettingsPanel.tsx"
    },
    output: {
        filename: "scripts/[name].js",
        libraryTarget: "amd"
    },
    externals: [
        {
            "q": true,
            "react": true,
            "react-dom": true
        },
        /^VSS\/.*/, /^TFS\/.*/, /^q$/
    ],
    resolve: {
        extensions: [".webpack.js", ".web.js", ".ts", ".tsx", ".js"],
        moduleExtensions: ["-loader"],
        alias: { 
            "OfficeFabric": path.resolve(__dirname, "node_modules/office-ui-fabric-react/lib"),
            "VSSUI": path.resolve(__dirname, "node_modules/vss-ui"),
            "VSTS_Extension_Widgets": path.resolve(__dirname, "node_modules/vsts-extension-react-widgets")
        }        
    },
    module: {
        rules: [
            {
                test: /\.tsx?$/,
                use: "ts-loader"
            },
            {
                test: /\.s?css$/,
                use: [
                    { loader: "style-loader" },
                    { loader: "css-loader" },
                    { loader: "sass-loader" }
                ]
            },
            {
                test: /\.(otf|eot|svg|ttf|woff|woff2|gif)(\?.+)?$/,
                use: "url-loader?limit=4096&name=[name].[ext]"
            }
        ]
    },
    plugins: [
        new webpack.DefinePlugin({
            "process.env.NODE_ENV": JSON.stringify("production")
        }),
        new UglifyJSPlugin({
            uglifyOptions: {
                output: {
                    comments: false,
                    beautify: false
                }
            }
        }),
        new CopyWebpackPlugin([
            { from: "./node_modules/vss-web-extension-sdk/lib/VSS.SDK.min.js", to: "3rdParty/VSS.SDK.min.js" },
            { from: "./node_modules/es6-promise/dist/es6-promise.min.js", to: "3rdParty/es6-promise.min.js" },
            
            { from: "./src/configs", to: "configs" },
            { from: "./src/images", to: "images" },
            { from: "./src/index.html", to: "index.html" },
            { from: "./src/vss-extension.json", to: "vss-extension.json" },
            { from: "./README.md", to: "README.md" }
        ])
    ]
}