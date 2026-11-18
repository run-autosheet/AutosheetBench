export function normalizedConnector(value) {
  const connector = String(value || "").trim().toLowerCase().replace(/[_\s]+/g, "-")
  if (connector === "google-sheets" || connector === "sheets") return "google-sheets"
  if (connector === "autosheet" || connector === "autosheets") return "autosheet"
  if (connector === "google-drive") return "google-drive"
  return null
}

export function autosheetProductLabel(item) {
  if (String(item?.connector || "").trim()) return null
  const text = [item?.source, item?.product, item?.label, item?.rawLabel, item?.id, item?.runConfigId]
    .filter(Boolean).join(" ")
  const legacy = /\bgpt[\s_-]*(?:for|4)[\s_-]*(?:work|excel|sheets)\b|\bgptforwork\b/i
  const native = /\bautosheet[\s_-]+(?:in|for)[\s_-]+(?:excel|sheets)\b/i
  if (!legacy.test(text) && !native.test(text)) return null
  const sheets = /sheets/i.test(item?.platform || item?.executionKind || "")
    || item?.surface === "sheets-addon"
    || /\b(?:gpt[\s_-]*(?:for|4)|autosheet[\s_-]+(?:in|for))[\s_-]+sheets\b/i.test(text)
  return sheets ? "Autosheet in Sheets" : "Autosheet in Excel"
}

export function connectorIdentity(item) {
  const connector = normalizedConnector(item?.connector)
  if (!connector) return null
  const identityText = [item?.source, item?.rawLabel, item?.label, item?.runConfigId, item?.id, item?.product]
    .filter(Boolean).join(" ").toLowerCase()
  const platform = identityText.includes("claude") ? "Claude Code"
    : identityText.includes("chatgpt")
      ? /chatgpt[\s_-]+work\b|\bwork-web\b/.test(identityText) || item?.surface === "work-web"
        ? "ChatGPT Work" : "ChatGPT"
      : null
  if (!platform) return null
  const label = connector === "google-drive" ? "Drive / Sheets MCP"
    : connector === "google-sheets" ? "Sheets MCP" : "Autosheet MCP"
  return { key: connector, label, platform }
}

export function connectorProductLabel(item) {
  const identity = connectorIdentity(item)
  return identity ? `${identity.platform} · ${identity.label}` : null
}

