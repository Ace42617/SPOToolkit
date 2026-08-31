/**
 * Guest / external-user detection for Permissions Matrix and remediation.
 *
 * SharePoint guest logins look like:
 *   i:0#.f|membership|user_contoso.com#ext#@tenant.onmicrosoft.com
 * Guest emails contain #EXT#@.
 *
 * Azure AD security groups use c:0t.c|tenant|{objectId} and are internal
 * principals — they must not be classified as external. Company Administrator
 * uses the same prefix.
 */

export function isExternalUser(loginName, email) {
  const emailStr = email == null ? "" : String(email);
  if (/#EXT#@/i.test(emailStr)) return true;
  const login = loginName == null ? "" : String(loginName);
  if (!login) return false;
  return /#ext#/i.test(login);
}
