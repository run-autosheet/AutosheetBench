function stableStringify(value) {
  if (Array.isArray(value)) {
    return `[${value.map((item) => stableStringify(item)).join(",")}]`
  }
  if (value && typeof value === "object") {
    return `{${Object.keys(value)
      .sort()
      .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
      .join(",")}}`
  }
  return JSON.stringify(value ?? null)
}

export function hashStable(value, length = 16) {
  const text = stableStringify(value)
  let hashA = 2166136261
  let hashB = 16777619
  for (let index = 0; index < text.length; index += 1) {
    const code = text.charCodeAt(index)
    hashA ^= code
    hashA = Math.imul(hashA, 16777619)
    hashB ^= code + index
    hashB = Math.imul(hashB, 2246822519)
  }
  const combined = [
    (hashA >>> 0).toString(16).padStart(8, "0"),
    (hashB >>> 0).toString(16).padStart(8, "0"),
    ((hashA ^ hashB) >>> 0).toString(16).padStart(8, "0"),
  ].join("")
  return combined.slice(0, length)
}

