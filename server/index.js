const express = require("express");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = process.env.PORT || 80;
const DIST_DIR = path.join(__dirname, "dist");

// Trust proxy for X-Forwarded-Proto header (needed when behind Traefik/nginx)
app.set("trust proxy", true);

// Read manifest template once at startup
const manifestTemplatePath = path.join(__dirname, "manifest.xml");
const manifestTemplate = fs.readFileSync(manifestTemplatePath, "utf-8");

// Dynamic manifest endpoint - MUST be before static files middleware
app.get("/manifest.xml", (req, res) => {
	// Explicitly check X-Forwarded-Proto for HTTPS behind Traefik
	let scheme = req.protocol;
	const xfp = req.get("X-Forwarded-Proto");
	if (xfp) {
		scheme = xfp.split(",")[0].trim();
	}
	const host = req.get("Host"); // 'localhost:3000' or 'officellm.apps.rp.alliance.co.uk'
	console.log(`[manifest] scheme=${scheme}, xfp=${xfp}, host=${host}`);
	const baseUrl = `${scheme}://${host}`;

	// Replace placeholders in manifest template
	const manifest = manifestTemplate
		.replace(/{{SCHEME}}/g, scheme)
		.replace(/{{HOST}}/g, host)
		.replace(/{{BASE_URL}}/g, baseUrl);

	res.set("Content-Type", "application/xml");
	res.send(manifest);
});

// Health check endpoint
app.get("/health", (_req, res) => {
	res.json({ status: "ok" });
});

// Serve static files from dist directory (after dynamic routes)
app.use(express.static(DIST_DIR));

app.listen(PORT, () => {
	console.log(`Office LLM Harness static server listening on port ${PORT}`);
});
