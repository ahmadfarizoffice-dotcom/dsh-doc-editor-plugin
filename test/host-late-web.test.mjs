/**
 * A fresh Harness boot applies plugins before the Web carrier is listening.
 * The Documents routes must wait for `webServer` instead of silently skipping
 * registration (which left every editor request at 404 until a hot reload).
 */
import assert from "node:assert/strict";
import { test } from "node:test";
import { mkdtemp, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";
import { fileURLToPath } from "node:url";

const here = path.dirname(fileURLToPath(import.meta.url));
const pluginDir = path.resolve(here, "..");
const plugin = await import(
  pathToFileURL(path.join(pluginDir, "index.js")).href
);

function makeServer() {
  const routes = [];
  return {
    routes,
    getOrigin: () => "http://127.0.0.1:1",
    register(route) {
      routes.push(route);
      return () => {
        const index = routes.indexOf(route);
        if (index >= 0) routes.splice(index, 1);
      };
    },
    registerFallback: () => () => {},
  };
}

test("routes register when the web server appears after activation", async () => {
  const workspace = await mkdtemp(path.join(tmpdir(), "dsh-late-web-"));
  const registeredTools = [];
  const effectDisposers = [];
  let pendingInject;
  const ctx = {
    logger: { warn() {}, info() {}, error: console.error },
    tools: {
      register: (definition) => (registeredTools.push(definition), () => {}),
    },
    effect(fn) {
      const dispose = fn();
      if (typeof dispose === "function") effectDisposers.push(dispose);
      return () => dispose?.();
    },
    get: (name) => (name === "webServer" ? undefined : undefined),
    inject(deps, callback) {
      assert.deepEqual(deps, ["webServer"]);
      pendingInject = callback;
    },
  };

  plugin.apply(ctx, { workspaceDir: workspace });
  assert.equal(
    typeof pendingInject,
    "function",
    "apply must wait for webServer",
  );
  assert.ok(
    registeredTools.length > 0,
    "tools still register without a web server",
  );

  const server = makeServer();
  const routeCtx = {
    get: (name) => (name === "webServer" ? server : undefined),
    effect(fn) {
      const dispose = fn();
      if (typeof dispose === "function") effectDisposers.push(dispose);
      return () => dispose?.();
    },
  };
  pendingInject(routeCtx);

  assert.ok(
    server.routes.length >= 1,
    "the route table is registered once webServer exists",
  );
  assert.equal(server.routes[0].path, "/api/dsh-documents");
  assert.equal(server.routes[0].kind, "prefix");

  // A second activation on the same route ctx must not double-register.
  const before = server.routes.length;
  pendingInject(routeCtx);
  assert.equal(
    server.routes.length,
    before * 2,
    "each activation owns one route row",
  );

  for (const dispose of effectDisposers.reverse()) {
    try {
      dispose();
    } catch {
      // best effort
    }
  }
  await rm(workspace, { recursive: true, force: true });
});

test("immediate web server keeps the synchronous registration path", async () => {
  const workspace = await mkdtemp(path.join(tmpdir(), "dsh-web-now-"));
  const server = makeServer();
  const ctx = {
    logger: { warn() {}, info() {}, error: console.error },
    tools: { register: () => () => {} },
    effect(fn) {
      fn();
      return () => {};
    },
    get: (name) => (name === "webServer" ? server : undefined),
  };
  plugin.apply(ctx, { workspaceDir: workspace });
  assert.equal(server.routes.length, 1);
  assert.equal(server.routes[0].path, "/api/dsh-documents");
  await rm(workspace, { recursive: true, force: true });
});
