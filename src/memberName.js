export function getMemberDisplayName(account = {}, username = "") {
  const rawName = (account.displayName || "").trim();
  if (rawName) return rawName;
  const loginName = (username || "").trim();
  return loginName || "Member";
}
