// Serves the add-in over HTTPS on https://localhost:3000 using the trusted
// Office dev certificate (run `npm run certs` once first).
import { spawn } from "node:child_process";
import { createRequire } from "node:module";

const require = createRequire(import.meta.url);
const devCerts = require("office-addin-dev-certs");

const options = await devCerts.getHttpsServerOptions();
const fs = await import("node:fs");
const os = await import("node:os");
const path = await import("node:path");

// http-server needs the cert and key as files.
const dir = fs.mkdtempSync(path.join(os.tmpdir(), "addin-certs-"));
const certFile = path.join(dir, "cert.pem");
const keyFile = path.join(dir, "key.pem");
fs.writeFileSync(certFile, options.cert);
fs.writeFileSync(keyFile, options.key);

const server = spawn(
  process.execPath,
  [require.resolve("http-server/bin/http-server"), ".", "-p", "3000", "-S", "-C", certFile, "-K", keyFile, "-c-1", "--cors"],
  { stdio: "inherit" }
);
server.on("exit", (code) => {
  fs.rmSync(dir, { recursive: true, force: true });
  process.exit(code ?? 0);
});
