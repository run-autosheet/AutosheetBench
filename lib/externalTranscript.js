const PRODUCT_PANE_ANCHORS = {
  chatgpt: /HTML content.*(?:Basispoints|ChatGPT)/i,
  claude: /HTML content.*Claude/i,
  copilot: /HTML content.*Copilot/i,
  gptforwork: /HTML content.*GPT/i,
}

const FINAL_ANSWER_MARKERS = {
  chatgpt: /\bbutton Copy response\b/i,
  claude: /\bbutton Copy message\b/i,
  copilot: /\btoolbar Message actions\b/i,
  gptforwork: /\b(?:(?:pop(?:-| )up )?button )?Copy (?:conversation|response|answer)\b/i,
}

const CHAT_NOISE = [
  /^(?:You said|The assistant said|Copilot said):?$/i,
  /^(?:ChatGPT|Claude|Copilot|GPT for Work)$/i,
  /^(?:Action|Chat|Edit|Agent)$/i,
  /^\|$/,
  /^\d{1,2}:\d{2}$/,
  /^New chat$/i,
  /^Scroll to bottom$/i,
  /^Copy (?:message|conversation|response|answer)(?:\s+⌘C)?$/i,
  /^Save to prompt library$/i,
  /^Edit this message/i,
  /^Fork chat/i,
  /^Restore workbook/i,
  /^Share$/i,
  /^Start dictation$/i,
  /^Disable sound notifications$/i,
  /^Suggested follow-ups$/i,
  /^Rate your experience$/i,
  /^AI-generated content may be incorrect$/i,
  /^Reply,? or \/skillify/i,
  /^Ask anything/i,
  /^Describe what you'd like to edit$/i,
  /^Current workbook selection:/i,
  /^\S+(?:\s+\S+)*\s+[A-Z]{1,3}\d+ selected$/i,
  /^Model (?:and thinking effort|selector):/i,
  /^\d+\s*% context(?: used)?$/i,
  /^Run time\s+\d+s$/i,
  /^(?:Docs|Demo|Ready)$/i,
  /^benchmark-live-[\w-]+$/i,
  /^(?:Bulk tools|GPT functions|Usage|Admin settings|What's new|Documentation|Contact support|Log out|History|Examples|New)$/i,
  /^Close GPT for Excel Word\b/i,
  /^Turn this into a reusable skill$/i,
  /^version\s+[\d.]+$/i,
  /^[\w.+-]+@[\w.-]+\.[A-Za-z]{2,}$/,
  /^GPT-[\d.]+\s+\w+$/i,
  /^(?:High|Medium|Low|Default)$/i,
  /^Add-in options\b/i,
  /^Failed\s*[—-]\s*the last run in this chat ended in an error$/i,
  /^Window:\s*"?output"?/i,
  /^\d+\s+standard window\b/i,
  /^\d+s$/i,
  /^Thought for\b/i,
  /^Reasoned in\b/i,
  /^(?:Inspecting|Reviewing|Scanning|Reading|Checking|Analy[sz]ing|Working on)\b.{0,140}$/i,
]

function normalizedKey(value) {
  return String(value || "").replace(/\s+/g, " ").trim().toLowerCase()
}

function indentationWidth(value) {
  let width = 0
  for (const character of String(value || "")) {
    width += character === "\t" ? 4 : 1
  }
  return width
}

function parseAccessibilityLine(rawLine, lineIndex) {
  const match = String(rawLine || "").match(/^(\s*)(\d+)\s+(.+)$/)
  if (!match) return null
  const body = match[3].trim()
  const roleMatch = body.match(/^(HTML content|text entry area|pop(?:-| )up button|toggle button|menu button|radio button|content list|list marker|heading|toolbar|button|container|row|cell|text|link)(?:\s+\([^)]*\))*\s*(.*)$/i)
  if (!roleMatch) return null
  return {
    lineIndex,
    indent: indentationWidth(match[1]),
    body,
    role: roleMatch ? roleMatch[1].toLowerCase() : "",
    content: roleMatch ? roleMatch[2].trim() : body,
  }
}

function productKey(product) {
  const key = String(product || "")
    .trim()
    .toLowerCase()
    .replace(/\s+in excel$/, "")
    .replace(/\s+/g, "")
  return key === "gp4w" ? "gptforwork" : key
}

// A formula chip exposes each of its tokens as a text record: "=", then
// "S", "U", "M", ... "(", "F2:F8", ",". Read one by one, the tokens repeat
// and removeRedundantLines drops the repeats, so the chip is joined back
// into its formula first: an "=" record and the token siblings that follow
// it at the same indent, none holding a space outside its quotes. Chips
// listed one after another, as in a table's formula column, follow on with
// no record between them, so an "=" outside any parentheses starts the next
// formula, after a token that ends an operand; one inside them, as in
// IF(H65=0, ...), or after an operator, as in a format's [>=1000], is a
// comparison.
function joinFormulaChips(records) {
  const unquoted = (text) => text.replace(/"[^"]*"|'[^']*'/g, "")
  const isToken = (record, indent) =>
    record?.role === "text" &&
    record.indent === indent &&
    record.content.length > 0 &&
    !/\s/.test(unquoted(record.content))
  const joined = []
  for (let index = 0; index < records.length; index += 1) {
    const record = records[index]
    let end = index + 1
    if (record.role === "text" && record.content === "=") {
      while (isToken(records[end], record.indent)) end += 1
    }
    if (end - index < 3) {
      joined.push(record)
      continue
    }
    const formulas = []
    let depth = 0
    for (const token of records.slice(index, end)) {
      const previous = formulas.at(-1)?.at(-1)?.content || ""
      if (
        token.content === "=" &&
        (!formulas.length || (depth === 0 && /[\w)"'$]$/.test(previous)))
      )
        formulas.push([])
      formulas.at(-1).push(token)
      const bare = unquoted(token.content)
      depth += (bare.match(/\(/g) || []).length - (bare.match(/\)/g) || []).length
    }
    for (const tokens of formulas) {
      const content = tokens.map((token) => token.content).join("")
      joined.push({ ...tokens[0], body: `text ${content}`, content })
    }
    index = end - 1
  }
  return joined
}

function paneRecords(rawText, product) {
  const records = joinFormulaChips(
    String(rawText || "")
      .split(/\r?\n/)
      .map(parseAccessibilityLine)
      .filter(Boolean),
  )
  const anchorPattern = PRODUCT_PANE_ANCHORS[productKey(product)]
  if (!anchorPattern) return records
  const anchorIndex = records.findIndex((record) => anchorPattern.test(record.body))
  if (anchorIndex < 0) return records
  const anchorIndent = records[anchorIndex].indent
  const paneTail = records.slice(anchorIndex + 1)
  // The GPT for Work runner's pane-only fallback intentionally returns a
  // flat numbered list rather than the full Excel accessibility hierarchy.
  // In that representation every record has the same indentation as the
  // HTML-content anchor, so treating the first peer as the pane boundary
  // produces an empty transcript and leaks the raw accessibility dump into
  // the Inspector. The capture already contains only the product pane; keep
  // its complete tail and let the message/noise filters below clean it.
  if (!paneTail.some((record) => record.indent > anchorIndent)) return paneTail
  let endIndex = records.length
  for (let index = anchorIndex + 1; index < records.length; index += 1) {
    if (records[index].indent <= anchorIndent) {
      endIndex = index
      break
    }
  }
  return records.slice(anchorIndex + 1, endIndex)
}

function stripAccessibilityMetadata(value) {
  return String(value || "")
    .replace(/[\u200B-\u200D\u2060\uFEFF\u00AD]/g, "")
    .replace(/,\s*(?:Help|Secondary Actions|Description|Value|ID):.*$/i, "")
    .replace(/^Value:\s*/i, "")
    .replace(/^[:\s—-]+/, "")
    .replace(/\s+/g, " ")
    .trim()
}

function cellReferenceFromButton(value) {
  const text = stripAccessibilityMetadata(value)
  const match = text.match(/^(?:Navigate to|Go to (?:cell|range|chart)|Select)\s+(.+)$/i)
  return match ? match[1].trim() : ""
}

function isNoise(value, promptKeys) {
  const text = stripAccessibilityMetadata(value)
  if (!text) return true
  const key = normalizedKey(text)
  if (promptKeys.has(key)) return true
  return CHAT_NOISE.some((pattern) => pattern.test(text))
}

function removeRedundantLines(lines) {
  const unique = []
  const seen = new Set()
  for (const line of lines) {
    const text = stripAccessibilityMetadata(line)
    const key = normalizedKey(text)
    // A formula the reasoning wrote may come again in the answer.
    if (!text || (seen.has(key) && !/^=\S/.test(text))) continue
    seen.add(key)
    unique.push(text)
  }
  return unique.filter((line, index) => {
    if (line.length < 20) return true
    const key = normalizedKey(line)
    return !unique.some((other, otherIndex) => {
      if (otherIndex === index || other.length <= line.length + 8) return false
      const otherKey = normalizedKey(other)
      if (!otherKey.includes(key)) return false
      // Accessibility trees expose a table row once whole and once per cell.
      // Drop the cell only when the longer line is mostly made of nearby
      // lines; a sentence that merely mentions it must not erase it.
      const covered = unique
        .slice(Math.max(0, otherIndex - 8), otherIndex + 9)
        .filter((part) => part !== other && otherKey.includes(normalizedKey(part)))
        .reduce((total, part) => total + normalizedKey(part).length, 0)
      return covered >= 0.6 * otherKey.length
    })
  })
}

function messageTextFromRecords(records, prompts) {
  const promptKeys = new Set((prompts || []).map(normalizedKey).filter(Boolean))
  const lines = []
  let skippedSubtreeIndent = null
  for (const record of records) {
    if (skippedSubtreeIndent !== null) {
      if (record.indent > skippedSubtreeIndent) continue
      skippedSubtreeIndent = null
    }
    if (record.role.includes("button") && /^(?:Thinking|Reasoned in)\b/i.test(stripAccessibilityMetadata(record.content))) {
      skippedSubtreeIndent = record.indent
      continue
    }
    if (record.role === "button") {
      const reference = cellReferenceFromButton(record.content)
      if (reference) lines.push(reference)
      continue
    }
    if (!["text", "heading", "link", "row", "cell"].includes(record.role)) continue
    const text = stripAccessibilityMetadata(record.content)
    if (isNoise(text, promptKeys)) continue
    // ChatGPT sometimes exposes one concatenated reasoning/status line before
    // the final response fragments. Prefer the final fragments that follow.
    if (/\bThought for\b/i.test(text)) continue
    lines.push(text)
  }
  return removeRedundantLines(lines).join("\n").trim()
}

function copilotVisibleTraceEntries(rawText, prompts = []) {
  const records = paneRecords(rawText, "copilot")
  const positions = promptPositions(records, prompts)
  const promptKeys = new Set((prompts || []).map(normalizedKey).filter(Boolean))
  const entries = []

  for (let turn = 0; turn < prompts.length; turn += 1) {
    entries.push({ role: "user", turn: turn + 1, text: prompts[turn] })
    const start = positions[turn] >= 0 ? positions[turn] + 1 : -1
    if (start < 0) continue
    const nextPrompt = positions.slice(turn + 1).find((position) => position >= 0)
    const end = nextPrompt === undefined ? records.length : nextPrompt
    const segment = records.slice(start, end)
    const reasoningRanges = []

    for (let index = 0; index < segment.length; index += 1) {
      const record = segment[index]
      if (!record.role.includes("button") || !/^Reasoned in\b/i.test(stripAccessibilityMetadata(record.content))) {
        continue
      }
      const reasoningLines = []
      let nextIndex = index + 1
      let detailIndent = record.indent
      if (
        segment[nextIndex]?.role === "container" &&
        /^\d+ steps? completed$/i.test(stripAccessibilityMetadata(segment[nextIndex].content))
      ) {
        detailIndent = segment[nextIndex].indent
        nextIndex += 1
      }
      while (nextIndex < segment.length && segment[nextIndex].indent > detailIndent) {
        const detail = segment[nextIndex]
        if (["text", "heading", "link"].includes(detail.role)) {
          const text = stripAccessibilityMetadata(detail.content).replace(/\s*\|\s*\d+\s*$/, "")
          if (!isNoise(text, promptKeys)) reasoningLines.push(text)
        }
        nextIndex += 1
      }
      reasoningRanges.push([index, nextIndex])
      const text = removeRedundantLines(reasoningLines).join("\n").trim()
      if (text) {
        entries.push({
          role: "assistant",
          kind: "reasoning",
          turn: turn + 1,
          text,
          final: false,
          captureSource: "accessibility-clean",
        })
      }
      index = nextIndex - 1
    }

    const finalMarker = FINAL_ANSWER_MARKERS.copilot
    const markerIndex = segment.map((record) => record.body).findLastIndex((body) => finalMarker.test(body))
    const contentRecords = (markerIndex >= 0 ? segment.slice(0, markerIndex) : segment).filter((record, index) => (
      !reasoningRanges.some(([rangeStart, rangeEnd]) => index >= rangeStart && index < rangeEnd)
    ))
    const finalText = messageTextFromRecords(contentRecords, prompts)
    if (finalText) {
      entries.push({
        role: "assistant",
        turn: turn + 1,
        text: finalText,
        final: markerIndex >= 0,
        captureSource: "accessibility-clean",
      })
    }
  }

  return entries
}

const CLAUDE_TRACE_CHROME = [
  /^(?:You said|The assistant said):?$/i,
  /^Claude$/i,
  /^New chat$/i,
  /^Scroll to bottom$/i,
  /^Copy message$/i,
  /^Save to prompt library$/i,
  /^Current workbook selection:/i,
  /^Model selector/i,
  /^Reply,? or \/skillify/i,
  /^Ask Claude for help,? or \/ for skills/i,
  /^Start dictation$/i,
  /^AI-generated content may be incorrect$/i,
]

const CLAUDE_NON_TOOL_BUTTONS = [
  /^(?:Thinking|Reasoned in)\b/i,
  /^More options$/i,
  /^Accept all edits$/i,
  /^Copy message$/i,
  /^Scroll to bottom$/i,
  /^Start dictation$/i,
]

function isClaudeTraceNoise(value, promptKeys) {
  const text = stripAccessibilityMetadata(value)
  if (!text || promptKeys.has(normalizedKey(text))) return true
  return CLAUDE_TRACE_CHROME.some((pattern) => pattern.test(text))
}

function claudeToolLabel(record) {
  if (record?.role !== "button" || !/^button\s+\((?:collapsed|expanded)\)/i.test(record.body)) return ""
  const label = stripAccessibilityMetadata(record.content)
  if (!label || CLAUDE_NON_TOOL_BUTTONS.some((pattern) => pattern.test(label))) return ""
  return label
}

function claudeVisibleTraceEntries(rawText, prompts = []) {
  const records = paneRecords(rawText, "claude")
  const positions = promptPositions(records, prompts)
  const promptKeys = new Set((prompts || []).map(normalizedKey).filter(Boolean))
  const entries = []

  for (let turn = 0; turn < prompts.length; turn += 1) {
    entries.push({ role: "user", turn: turn + 1, text: prompts[turn] })
    const start = positions[turn] >= 0 ? positions[turn] + 1 : -1
    if (start < 0) continue
    const nextPrompt = positions.slice(turn + 1).find((position) => position >= 0)
    const end = nextPrompt === undefined ? records.length : nextPrompt
    const segment = records.slice(start, end)
    const assistantLines = []
    let toolIndex = 0
    let sawFinalMarker = false

    const flushAssistant = () => {
      const text = removeRedundantLines(assistantLines.splice(0)).join("\n").trim()
      if (text) {
        entries.push({
          role: "assistant",
          turn: turn + 1,
          text,
          final: false,
          captureSource: "accessibility-clean",
        })
      }
    }

    for (let index = 0; index < segment.length; index += 1) {
      const record = segment[index]
      if (FINAL_ANSWER_MARKERS.claude.test(record.body)) {
        sawFinalMarker = true
        continue
      }
      const toolLabel = claudeToolLabel(record)
      if (toolLabel) {
        flushAssistant()
        const detailLines = []
        let nextIndex = index + 1
        while (nextIndex < segment.length && segment[nextIndex].indent > record.indent) {
          const detail = segment[nextIndex]
          if (["text", "heading", "link"].includes(detail.role)) {
            const detailText = stripAccessibilityMetadata(detail.content)
            if (!isClaudeTraceNoise(detailText, promptKeys)) detailLines.push(detailText)
          }
          nextIndex += 1
        }
        toolIndex += 1
        const details = removeRedundantLines(detailLines).join("\n").trim()
        entries.push({
          role: "tool_call",
          turn: turn + 1,
          callId: `claude-ui-${turn + 1}-${toolIndex}`,
          toolName: toolLabel,
          status: "visible in UI",
          argsPreview: details ? "Visible details" : "Visible Claude action",
          argsFull: details,
          captureSource: "accessibility-clean",
        })
        index = nextIndex - 1
        continue
      }
      if (!["text", "heading", "link"].includes(record.role)) continue
      const text = stripAccessibilityMetadata(record.content)
      if (!isClaudeTraceNoise(text, promptKeys)) assistantLines.push(text)
    }
    flushAssistant()
    if (sawFinalMarker) {
      const lastAssistant = entries.findLast((entry) => entry.role === "assistant" && entry.turn === turn + 1)
      if (lastAssistant) lastAssistant.final = true
    }
  }

  return entries
}

function claudeNativeText(value) {
  if (typeof value === "string") return value.trim()
  if (Array.isArray(value)) {
    return value
      .map((item) => claudeNativeText(item))
      .filter(Boolean)
      .join("\n")
      .trim()
  }
  if (value && typeof value === "object") {
    if (typeof value.text === "string") return value.text.trim()
    try {
      return JSON.stringify(value, null, 2)
    } catch {
      return String(value)
    }
  }
  return value == null ? "" : String(value).trim()
}

function isClaudeNativeContextBlock(value) {
  return /^<(?:system-reminder|user_context|initial_state)>/i.test(String(value || "").trim()) ||
    /^\[id:[^\]]+\]$/i.test(String(value || "").trim())
}

export function claudeNativeTranscriptEntries(nativeTranscript) {
  const messages = Array.isArray(nativeTranscript?.messages) ? nativeTranscript.messages : []
  if (messages.length === 0) return []

  const entries = []
  const toolNames = new Map()
  let turn = 0

  for (const message of messages) {
    const role = String(message?.role || "").toLowerCase()
    const blocks = Array.isArray(message?.content) ? message.content : []
    if (role === "user") {
      const authoredText = blocks
        .filter((block) => block?.type === "text")
        .map((block) => claudeNativeText(block?.text))
        .filter((text) => text && !isClaudeNativeContextBlock(text))
      if (authoredText.length > 0) {
        turn += 1
        entries.push({
          role: "user",
          turn,
          text: authoredText.join("\n"),
          captureSource: "claude-native-export",
        })
      }
      for (const block of blocks.filter((item) => item?.type === "tool_result")) {
        const callId = String(block?.tool_use_id || "")
        entries.push({
          role: "tool_result",
          turn: Math.max(1, turn),
          callId,
          toolName: toolNames.get(callId) || "tool",
          status: block?.is_error === true ? "error" : "success",
          resultPreview: block?.is_error === true ? "Error" : "Result",
          resultFull: claudeNativeText(block?.content),
          captureSource: "claude-native-export",
        })
      }
      continue
    }
    if (role !== "assistant") continue

    const messageTurn = Math.max(1, turn)
    const hasToolUse = blocks.some((block) => block?.type === "tool_use")
    const assistantEntries = []
    for (const block of blocks) {
      if (block?.type === "tool_use") {
        const callId = String(block?.id || "")
        const toolName = String(block?.name || "tool")
        if (callId) toolNames.set(callId, toolName)
        entries.push({
          role: "tool_call",
          turn: messageTurn,
          callId,
          toolName,
          status: "called",
          argsPreview: "Arguments",
          argsFull: claudeNativeText(block?.input),
          captureSource: "claude-native-export",
        })
        continue
      }
      if (block?.type !== "text" && block?.type !== "thinking") continue
      const text = claudeNativeText(block?.text ?? block?.thinking)
      if (!text) continue
      const entry = {
        role: "assistant",
        turn: messageTurn,
        text,
        final: false,
        captureSource: "claude-native-export",
      }
      if (block?.type === "thinking") entry.kind = "reasoning"
      entries.push(entry)
      assistantEntries.push(entry)
    }
    if (!hasToolUse && assistantEntries.length > 0) {
      assistantEntries[assistantEntries.length - 1].final = true
    }
  }

  return entries
}

export function chatGptNativeMarkdownEntries(markdown) {
  const source = String(markdown || "").replace(/\r\n?/g, "\n").trim()
  if (!/^# Conversation export\b/m.test(source)) return []

  const sectionPattern = /^## (User|Assistant reasoning|Assistant|Tool call)\s*$/gm
  const sections = []
  let match
  while ((match = sectionPattern.exec(source)) !== null) {
    sections.push({ kind: match[1], start: match.index, contentStart: sectionPattern.lastIndex })
  }
  if (sections.length === 0) return []

  const entries = []
  let turn = 0
  let toolIndex = 0
  for (let index = 0; index < sections.length; index += 1) {
    const section = sections[index]
    const end = sections[index + 1]?.start ?? source.length
    let body = source.slice(section.contentStart, end).trim()
    let timestamp = null
    const timestampMatch = body.match(/^_([^\n]+)_\s*(?:\n|$)/)
    if (timestampMatch) {
      timestamp = timestampMatch[1].trim()
      body = body.slice(timestampMatch[0].length).trim()
    }
    if (!body) continue

    if (section.kind === "User") {
      turn += 1
      entries.push({
        role: "user",
        turn,
        text: body,
        timestamp,
        captureSource: "chatgpt-native-export",
      })
      continue
    }
    if (section.kind === "Tool call") {
      toolIndex += 1
      const toolName = body.match(/^- Tool:\s*(.+)$/mi)?.[1]?.trim() || "tool"
      const nativeStatus = body.match(/^- Status:\s*(.+)$/mi)?.[1]?.trim() || ""
      entries.push({
        role: "tool_call",
        turn: Math.max(1, turn),
        callId: `chatgpt-native-tool-${toolIndex}`,
        toolName,
        status: /(?:complete|success)/i.test(nativeStatus)
          ? "success"
          : /(?:error|fail)/i.test(nativeStatus)
            ? "error"
            : "called",
        argsPreview: nativeStatus || "Tool call",
        argsFull: body,
        timestamp,
        captureSource: "chatgpt-native-export",
      })
      continue
    }
    entries.push({
      role: "assistant",
      turn: Math.max(1, turn),
      text: body,
      kind: section.kind === "Assistant reasoning" ? "reasoning" : undefined,
      final: false,
      timestamp,
      captureSource: "chatgpt-native-export",
    })
  }

  const finalAssistant = entries.findLast((entry) => (
    entry.role === "assistant" && entry.kind !== "reasoning"
  ))
  if (finalAssistant) finalAssistant.final = true
  return entries
}

function promptPositions(records, prompts) {
  const promptList = prompts || []
  const occurrences = promptList.map((prompt) => {
    const key = normalizedKey(prompt)
    return records
      .map((record, index) => normalizedKey(stripAccessibilityMetadata(record.content)) === key ? index : -1)
      .filter((index) => index >= 0)
  })
  if (occurrences.length === 0) return []

  // Accessibility snapshots include the add-in's history before the active
  // conversation. Select the last ordered prompt sequence so history is not
  // mistaken for the current turn.
  let best = null
  function visit(turn, previous, selected) {
    if (turn >= occurrences.length) {
      if (!best || selected[selected.length - 1] > best[best.length - 1]) best = selected
      return
    }
    for (const position of occurrences[turn]) {
      if (position <= previous) continue
      visit(turn + 1, position, [...selected, position])
    }
  }
  visit(0, -1, [])
  return best || occurrences.map((items) => items[0] ?? -1)
}

export function externalFinalAnswerCount(rawText, product) {
  const marker = FINAL_ANSWER_MARKERS[productKey(product)]
  if (!marker) return 0
  return String(rawText || "")
    .split(/\r?\n/)
    .filter((line) => marker.test(line))
    .length
}

export function extractExternalAssistantMessages(rawText, product, prompts = []) {
  const records = paneRecords(rawText, product)
  if (records.length === 0) {
    const promptKeys = new Set((prompts || []).map(normalizedKey).filter(Boolean))
    const rawLines = String(rawText || "").split(/\r?\n/)
    // Some JXA fallbacks return a flat list of accessible labels instead of
    // numbered role records. The same pane contains chat history, so start at
    // the last occurrence of the active prompt rather than including history.
    let activeStart = -1
    for (let index = 0; index < rawLines.length; index += 1) {
      if (promptKeys.has(normalizedKey(stripAccessibilityMetadata(rawLines[index])))) activeStart = index
    }
    const activeLines = activeStart >= 0 ? rawLines.slice(activeStart + 1) : rawLines
    const text = removeRedundantLines(
      activeLines
        .map(stripAccessibilityMetadata)
        .filter((line) => !isNoise(line, promptKeys)),
    ).join("\n").trim()
    return text
      ? [{
          turn: Math.max(1, prompts.length),
          text,
          final: externalFinalAnswerCount(rawText, product) > 0,
        }]
      : []
  }
  const positions = promptPositions(records, prompts)
  const finalMarker = FINAL_ANSWER_MARKERS[productKey(product)]
  const messages = []
  for (let turn = 0; turn < Math.max(1, prompts.length); turn += 1) {
    const start = positions[turn] >= 0 ? positions[turn] + 1 : (turn === 0 ? 0 : -1)
    if (start < 0) continue
    const nextPrompt = positions.slice(turn + 1).find((position) => position >= 0)
    const end = nextPrompt === undefined ? records.length : nextPrompt
    const segment = records.slice(start, end)
    const markerIndex = finalMarker
      ? segment.map((record) => record.body).findLastIndex((body) => finalMarker.test(body))
      : -1
    let contentRecords = markerIndex >= 0 ? segment.slice(0, markerIndex) : segment
    const text = messageTextFromRecords(contentRecords, prompts)
    if (text) messages.push({ turn: turn + 1, text, final: markerIndex >= 0 })
  }
  return messages
}

// Gemini in Sheets is captured as its panel's visible text, one line per
// element: the panel's controls and the prompt, then after "Gemini response"
// the model's thoughts (a title, then a paragraph in the first person), the
// plan of an edit ending with "Summary", the reply with a "Select" chip after
// each cell it names and the buttons of any chart it inserted, then the
// rating buttons and suggested questions. Split the thoughts from the reply,
// so the judge and the inspector read the reply as the final answer.
const GEMINI_CONTROLS = new Set(["Insert", "Preview", "Copy", "Retry", "More options"])

// A thought is written in the first person with contractions ("I'm now",
// "I've just"), where the reply says "I have added". The thoughts end at the
// last such paragraph; lead-in lines between them ("Here's the response:")
// go with them.
function isGeminiThought(line) {
  return line.length >= 40 && /\b(?:I'm|I've|I'll|I'd)\b/.test(line)
}

function geminiPanelEntries(entry) {
  let lines = String(entry?.text || "").split("\n").map((line) => line.trim())
  const start = lines.lastIndexOf("Gemini response")
  if (start < 0) return [entry]
  lines = lines.slice(start + 1)
  const end = lines.indexOf("Good response")
  if (end >= 0) lines = lines.slice(0, end)
  lines = lines.filter((line) => !GEMINI_CONTROLS.has(line))
  if (lines[0] === "Model thoughts") lines = lines.slice(1)
  const last = lines.findLastIndex(isGeminiThought)
  const thoughts = lines.slice(0, last + 1).filter(Boolean)
  lines = lines.slice(last + 1)
  const summary = lines.findIndex((line, index) => line === "Summary" && !lines[index + 1])
  const plan = summary >= 0 ? lines.slice(0, summary).filter(Boolean) : []
  if (summary >= 0) lines = lines.slice(summary + 1)
  const reply = lines
    .reduce((kept, line, index) => {
      if (/^Select \S+$/.test(line)) return kept
      if (/^Select \S+$/.test(lines[index - 1] || "") && kept.length) {
        kept[kept.length - 1] += /^[),.;:]/.test(line) ? line : ` ${line}`
      } else {
        kept.push(line)
      }
      return kept
    }, [])
    .filter(Boolean)
    .join("\n\n")
  const reasoning = [...thoughts, ...plan].join("\n")
  return [
    ...(reasoning ? [{ ...entry, text: reasoning, kind: "reasoning", final: false }] : []),
    ...(reply ? [{ ...entry, text: reply, final: true }] : []),
  ]
}

export function normalizeExternalTranscriptEntries(transcript, fallbackPrompts = []) {
  if (transcript?.captureMethod === 'visible-dom' && Array.isArray(transcript.prompts)) {
    return transcript.prompts.flatMap((prompt, index) => {
      const capture = transcript.turnResponses?.[index]?.visibleDom || (index === transcript.prompts.length - 1 ? transcript.visibleDom : '');
      const main = String(capture || '').split(/^- main:/m)[1] || '';
      const marker = main.lastIndexOf('heading "Claude a répondu');
      const response = marker >= 0 ? main.slice(marker).split(/\n\s*- toolbar "Actions du message"/)[0] : '';
      const lines = response.split('\n').slice(1).flatMap(line => {
        const match = line.match(/^\s*- (?:paragraph|text|code|strong|emphasis|row)(?:\s*:\s*|\s+)(.+?)\s*$/);
        if (!match) return [];
        let text = match[1].replace(/:\s*$/, '');
        if (text.startsWith('"') && text.endsWith('"')) { try { text = JSON.parse(text); } catch { text = text.slice(1,-1); } }
        return [text];
      });
      const text = lines.join('\n').trim() || response.match(/Claude a répondu : ([^\n]+?)"/)?.[1] || '';
      return [{ role: 'user', text: String(prompt), turn: index + 1 }, ...(text ? [{role:'assistant',text,turn:index+1}] : [])];
    });
  }
  const nativeClaudeTranscript = transcript?.nativeTranscript || (
    transcript?.meta && Array.isArray(transcript?.messages) ? transcript : null
  )
  const nativeClaudeEntries = claudeNativeTranscriptEntries(nativeClaudeTranscript)
  if (nativeClaudeEntries.length > 0) return nativeClaudeEntries
  const nativeChatGptEntries = chatGptNativeMarkdownEntries(
    transcript?.nativeTranscriptMarkdown || (
      typeof transcript?.nativeTranscript === "string" ? transcript.nativeTranscript : ""
    ),
  )
  if (nativeChatGptEntries.length > 0) return nativeChatGptEntries
  const entries = Array.isArray(transcript?.entries)
    ? transcript.entries
    : Array.isArray(transcript)
      ? transcript
      : []
  const recordedPrompts = entries
    .filter((entry) => entry?.role === "user" && normalizedKey(entry?.text) !== "see prompts.json")
    .map((entry) => String(entry.text || "").trim())
    .filter(Boolean)
  const prompts = recordedPrompts.length > 0
    ? recordedPrompts
    : (fallbackPrompts || []).map((prompt) => String(prompt || "").trim()).filter(Boolean)
  if (/^gemini/.test(productKey(transcript?.product))) {
    return entries.flatMap((entry) => entry?.role === "assistant" ? geminiPanelEntries(entry) : [entry])
  }
  if (productKey(transcript?.product) === "claude" && transcript?.rawAccessibilityText) {
    return claudeVisibleTraceEntries(transcript.rawAccessibilityText, prompts)
  }
  const rawAssistant = entries.find((entry) => (
    entry?.role === "assistant" && /^accessibility-(?:tree|tail)$/.test(entry?.captureSource || "")
  ))
  if (productKey(transcript?.product) === "copilot" && rawAssistant) {
    const visibleEntries = copilotVisibleTraceEntries(rawAssistant.text, prompts)
    const copiedFinalText = String(transcript?.copiedFinalText || "").trim()
    if (copiedFinalText) {
      const finalAssistant = visibleEntries.findLast((entry) => (
        entry.role === "assistant" && entry.kind !== "reasoning"
      ))
      if (finalAssistant) {
        finalAssistant.text = copiedFinalText
        finalAssistant.final = true
        finalAssistant.captureSource = "copy-control"
      } else {
        visibleEntries.push({
          role: "assistant",
          turn: Math.max(1, prompts.length),
          text: copiedFinalText,
          final: true,
          captureSource: "copy-control",
        })
      }
    }
    if (visibleEntries.some((entry) => entry.role === "assistant")) return visibleEntries
  }
  if (!rawAssistant) {
    return entries.map((entry) => {
      if (entry?.role !== "user" || normalizedKey(entry?.text) !== "see prompts.json") return entry
      return prompts.length === 1 ? { ...entry, text: prompts[0] } : entry
    })
  }
  const messages = extractExternalAssistantMessages(rawAssistant.text, transcript?.product, prompts)
  const normalized = []
  for (let index = 0; index < prompts.length; index += 1) {
    normalized.push({ role: "user", turn: index + 1, text: prompts[index] })
    const message = messages.find((item) => item.turn === index + 1)
    if (message) {
      normalized.push({
        role: "assistant",
        turn: index + 1,
        text: message.text,
        final: message.final,
        captureSource: "accessibility-clean",
      })
    }
  }
  return normalized
}

