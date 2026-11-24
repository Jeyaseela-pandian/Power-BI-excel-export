/* Sign-in for any organisation (multi-tenant): personal Microsoft accounts are told Power BI
   needs a work account, and organisations whose IT admin must approve apps get a clear
   message and the approval link to send them */
const test = require("node:test");
const assert = require("node:assert");
const { loadApp } = require("./harness");

const PERSONAL = "9188040d-6c67-4c5b-b112-36a304b66dad";

function app(config = {}) {
  const a = loadApp();
  Object.assign(a.ctx.window.APP_CONFIG, { clientId: "client-1", redirectUri: "http://localhost:5500", tenantId: "common" }, config);
  return a;
}

test("Sign-in: personal Microsoft accounts are recognised, work accounts aren't", () => {
  const { ctx } = app();
  assert.strictEqual(ctx.isPersonalAccount({ tenantId: PERSONAL }), true);
  assert.strictEqual(ctx.isPersonalAccount({ idTokenClaims: { tid: PERSONAL } }), true);
  assert.strictEqual(ctx.isPersonalAccount({ tenantId: "38d5e40c-9bd6-41fb-8951-c387c97018d9" }), false);
  assert.strictEqual(ctx.isPersonalAccount(null), false);
});

test("Sign-in: Microsoft's errors become messages a user can act on", () => {
  const { ctx, run } = app();
  const personal = run("PERSONAL_ACCOUNT_MESSAGE");
  assert.strictEqual(ctx.friendlyError(new Error("AADSTS500200: User account is a personal Microsoft account.")), personal);
  assert.match(ctx.friendlyError(new Error("AADSTS65004: User declined to consent")), /didn't approve/);
  assert.match(ctx.friendlyError(new Error("AADSTS90094: The grant requires admin permission")), /IT admin needs to approve/);
  assert.match(ctx.friendlyError(new Error("AADSTS65001: The user or administrator has not consented")), /IT admin needs to approve/);
  // The messages that were there before are unchanged
  assert.match(ctx.friendlyError(new Error("HTTP 401")), /session has expired/);
  assert.match(ctx.friendlyError(new Error("HTTP 404")), /wasn't found/);
});

test("Sign-in: the IT admin's approval link names this app and returns to it", () => {
  const url = new URL(app().ctx.adminConsentUrl());
  assert.strictEqual(url.origin, "https://login.microsoftonline.com");
  assert.strictEqual(url.pathname, "/organizations/adminconsent", "any organisation's admin can open it");
  assert.strictEqual(url.searchParams.get("client_id"), "client-1");
  assert.strictEqual(url.searchParams.get("redirect_uri"), "http://localhost:5500");
  // An app set up for one organisation sends its admin straight there
  const one = new URL(app({ tenantId: "38d5e40c-9bd6-41fb-8951-c387c97018d9" }).ctx.adminConsentUrl());
  assert.strictEqual(one.pathname, "/38d5e40c-9bd6-41fb-8951-c387c97018d9/adminconsent");
});

test("Sign-in: the answer to an admin approval is read from the address, other addresses are left alone", () => {
  const { json } = app();
  assert.deepStrictEqual(json(`adminConsentResult("?admin_consent=True&tenant=t1")`), { ok: true });
  assert.deepStrictEqual(json(`adminConsentResult("?error=access_denied&error_description=declined&tenant=t1")`), { ok: false, detail: "declined" });
  assert.strictEqual(json(`adminConsentResult("?reportId=r1&groupId=g1&pageName=p1")`), null, "report links are unaffected");
  assert.strictEqual(json(`adminConsentResult("")`), null);
});
