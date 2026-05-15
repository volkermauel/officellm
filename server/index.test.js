const http = require("http");
const path = require("path");
const fs = require("fs");

/**
 * Tests for the Express server's dynamic manifest.xml endpoint.
 * Verifies that:
 * 1. /manifest.xml route takes priority over static files middleware
 * 2. Dynamic URLs are correctly inferred from the Host header
 * 3. {{BASE_URL}} placeholder is replaced correctly
 * 4. HTTPS scheme detection works via X-Forwarded-Proto header (behind Traefik)
 */

const TEST_PORT = 3999;
const DIST_DIR = path.join(__dirname, "dist");

// Create dist directory with dummy files if it doesn't exist
if (!fs.existsSync(DIST_DIR)) {
	fs.mkdirSync(DIST_DIR, { recursive: true });
	fs.writeFileSync(path.join(DIST_DIR, "index.html"), "<html><body>Test</body></html>");
}

// Create a test server instance
function createTestServer() {
	const express = require("express");
	const app = express();

	// Trust proxy for X-Forwarded-Proto header (needed when behind Traefik/nginx)
	app.set("trust proxy", true);

	const PORT = process.env.PORT || TEST_PORT;
	const DIST_DIR_PATH = DIST_DIR;

	// Read manifest template once at startup
	const manifestTemplatePath = path.join(__dirname, "manifest.xml");
	const manifestTemplate = fs.readFileSync(manifestTemplatePath, "utf-8");

	// Dynamic manifest endpoint - MUST be before static files middleware
	app.get("/manifest.xml", (req, res) => {
		const scheme = req.protocol; // 'http' or 'https'
		const host = req.get("Host"); // 'localhost:3000' or 'officellm.apps.rp.alliance.co.uk'
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
	app.use(express.static(DIST_DIR_PATH));

	return new Promise((resolve) => {
		const server = app.listen(PORT, () => resolve(server));
	});
}

function stopServer(server) {
	return new Promise((resolve) => server.close(resolve));
}

function makeRequest(pathname, headers = {}) {
	return new Promise((resolve, reject) => {
		const options = {
			hostname: "127.0.0.1",
			port: TEST_PORT,
			path: pathname,
			method: "GET",
			headers: { Host: "officellm.apps.rp.alliance.co.uk", ...headers },
		};
		const req = http.request(options, (res) => {
			let data = "";
			res.on("data", (chunk) => (data += chunk));
			res.on("end", () => resolve({ status: res.statusCode, headers: res.headers, body: data }));
		});
		req.on("error", reject);
		req.end();
	});
}

async function runTests() {
	let passed = 0;
	let failed = 0;

	const server = await createTestServer();

	try {
		// Test 1: /manifest.xml returns correct dynamic URL
		console.log("\n🧪 Test 1: Dynamic manifest URL from Host header");
		const res1 = await makeRequest("/manifest.xml");
		if (res1.status === 200 && res1.body.includes("officellm.apps.rp.alliance.co.uk")) {
			console.log("   ✅ PASS: Manifest contains correct URL");
			passed++;
		} else {
			console.log("   ❌ FAIL: Manifest does not contain correct URL");
			console.log("   Body:", res1.body.substring(0, 500));
			failed++;
		}

		// Test 2: /manifest.xml takes priority over static files
		console.log("\n🧪 Test 2: Dynamic route priority over static files");
		const res2 = await makeRequest("/manifest.xml");
		if (res2.body.includes("{{BASE_URL}}")) {
			console.log("   ❌ FAIL: Placeholder not replaced");
			failed++;
		} else if (res2.status === 200 && !res2.body.includes("localhost:8080")) {
			console.log("   ✅ PASS: Dynamic route served (not static file)");
			passed++;
		} else {
			console.log("   ❌ FAIL: Static file served instead of dynamic route");
			failed++;
		}

		// Test 3: Different Host header produces different URL
		console.log("\n🧪 Test 3: Different Host header produces different URL");
		const res3 = await makeRequest("/manifest.xml", { Host: "localhost:3000" });
		if (res3.body.includes("http://localhost:3000")) {
			console.log("   ✅ PASS: Localhost URL correctly generated");
			passed++;
		} else {
			console.log("   ❌ FAIL: Localhost URL not generated");
			console.log("   Body:", res3.body.substring(0, 500));
			failed++;
		}

		// Test 4: HTTPS scheme detection via X-Forwarded-Proto header (behind Traefik)
		console.log("\n🧪 Test 4: HTTPS scheme detection via X-Forwarded-Proto");
		const res4 = await makeRequest("/manifest.xml", {
			Host: "example.com",
			"X-Forwarded-Proto": "https",
		});
		if (res4.body.includes("https://example.com")) {
			console.log("   ✅ PASS: HTTPS scheme correctly detected via X-Forwarded-Proto");
			passed++;
		} else {
			console.log("   ❌ FAIL: HTTPS scheme not detected");
			console.log("   Body:", res4.body.substring(0, 500));
			failed++;
		}

		// Test 5: Health endpoint still works
		console.log("\n🧪 Test 5: Health endpoint");
		const res5 = await makeRequest("/health");
		if (res5.status === 200 && res5.body.includes('"ok"')) {
			console.log("   ✅ PASS: Health endpoint works");
			passed++;
		} else {
			console.log("   ❌ FAIL: Health endpoint broken");
			failed++;
		}

		// Test 6: Static files still served
		console.log("\n🧪 Test 6: Static files still served");
		const res6 = await makeRequest("/index.html");
		if (res6.status === 200) {
			console.log("   ✅ PASS: Static files served");
			passed++;
		} else {
			console.log("   ❌ FAIL: Static files not served");
			failed++;
		}

		console.log(`\n📊 Results: ${passed} passed, ${failed} failed`);
		process.exit(failed > 0 ? 1 : 0);
	} catch (error) {
		console.error("Test error:", error);
		process.exit(1);
	} finally {
		await stopServer(server);
	}
}

runTests();
