const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = {
	entry: "./src/app.ts",
	output: {
		filename: "bundle.js",
		path: path.resolve(__dirname, "dist"),
		clean: true,
	},
	devtool: "source-map",
	mode: "development",
	module: {
		rules: [
			{
				test: /\.ts$/,
				use: "ts-loader",
				exclude: /node_modules/,
			},
		],
	},
	resolve: {
		extensions: [".ts", ".js"],
	},
	plugins: [
		new HtmlWebpackPlugin({
			template: "./src/index.html",
		}),
		new CopyWebpackPlugin({
			patterns: [
				{
					from: "manifest.xml",
					to: "../manifest.xml",
				},
			],
		}),
	],
	devServer: {
		port: 3000,
		host: "127.0.0.1",
		hot: true,
		// Office Add-ins need HTTPS in production; localhost HTTP works for dev
		// In production, use a reverse proxy with valid TLS
	},
};
