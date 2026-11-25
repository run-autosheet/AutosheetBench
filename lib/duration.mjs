function finiteDuration(value) {
  return typeof value === "number" && Number.isFinite(value) && value >= 0 ? value : null
}

// Recording measurements that run from prompt submission to the end of the
// answer, summed over turns. Measurements that start at the recording's start
// (recording-final-message-completion, visual-completion-backward-from-final-
// frame), runner observations and transcript timestamps are not admitted.
export const PROMPT_TO_END_RECORDING_MEASUREMENTS = new Set([
  "recording-submit-to-finish",
  "recording-working-label-sum-of-turns",
  "recording-stop-button-sum-of-turns",
  "recording-send-to-terminal-sum-of-turns",
  "recording-runner-submit-to-busy-label-end-sum-of-turns",
  "recording-stop-control-minus-question-wait",
  "recording-submit-to-feedback-controls",
])

const seconds = (value) => {
  const number = finiteDuration(value)
  return number === null ? null : number * 1000
}

// The one duration the website and the release show: prompt submission to
// the end of the answer, summed over turns.
//  1. the product's own agent time: Autosheet's backend agent time, Claude's
//     usage-dashboard session time, or the product's telemetry (prompt
//     submission to its last model response);
//  2. otherwise a recording measurement from prompt submission to the end;
//  3. otherwise, for a run stopped at its time limit, that limit;
//  4. otherwise the runner's clock, labelled runner-clock-upper-bound.
// `source` says which one, so pages can mark upper bounds.
// `permissionWait` ({seconds, publishedSeconds}, lib/permissionWaits.mjs)
// is the time an add-in permission card (Claude's or ChatGPT's) waited for the runner's click within
// that duration. It is subtracted when the duration is still the one it was
// measured against, and returned as `permissionWaitSeconds`. A timed-out run
// keeps its time limit.
export function promptToEndDuration(record, { permissionWait = null } = {}) {
  const duration = promptToEndDurationBeforeWaits(record)
  if (!duration || !permissionWait || duration.source === "time-limit") return duration
  if (record?.timedOut === true || record?.status === "timeout") return duration
  const waitMs = seconds(permissionWait.seconds), publishedMs = seconds(permissionWait.publishedSeconds)
  if (waitMs === null || publishedMs === null || Math.abs(duration.ms - publishedMs) > 10 || waitMs > duration.ms) return duration
  return { ...duration, ms: duration.ms - waitMs, permissionWaitSeconds: permissionWait.seconds }
}

function promptToEndDurationBeforeWaits(record) {
  const check = record?.timingCheck || {}
  const limitMs = seconds(record?.timeLimitSeconds)
  const timedOut = record?.timedOut === true || record?.status === "timeout"
  const capped = (ms) => (timedOut && limitMs !== null && ms > limitMs ? limitMs : ms)
  const runnerMs = finiteDuration(record?.totalElapsedMs) ?? finiteDuration(record?.durationMs) ?? seconds(record?.durationSeconds)

  const agentMs = finiteDuration(record?.agentElapsedMs) ?? finiteDuration(record?.agentDurationMs) ?? seconds(record?.agentDurationSeconds)
  if (agentMs !== null) {
    // The backend agent can keep counting after the runner stops a timed-out run.
    const ms = timedOut && runnerMs !== null && runnerMs < agentMs ? runnerMs : agentMs
    return { ms, source: "agent-time" }
  }
  const sessionMs = seconds(check.sessionDurationSeconds)
  if (sessionMs !== null) return { ms: capped(sessionMs), source: "claude-dashboard-session" }
  const telemetryMs = seconds(check.telemetryLastResponseSeconds)
  if (telemetryMs !== null) return { ms: capped(telemetryMs), source: "product-telemetry" }

  const recordedMs = seconds(check.recordingPromptToCompletionSeconds) ?? seconds(record?.recordingSubmitToFinishSeconds)
  if (recordedMs !== null) {
    // A recording whose boundary falls in a capture gap is measured to within a stated bound.
    const boundSeconds = finiteDuration(check.recordingApproximateBoundSeconds)
    return boundSeconds === null
      ? { ms: capped(recordedMs), source: "recording-prompt-to-end" }
      : { ms: capped(recordedMs), source: "recording-prompt-to-end-approximate", boundSeconds }
  }
  if (PROMPT_TO_END_RECORDING_MEASUREMENTS.has(record?.durationMeasurement) && runnerMs !== null) {
    return { ms: capped(runnerMs), source: "recording-prompt-to-end" }
  }
  if (timedOut && limitMs !== null) return { ms: limitMs, source: "time-limit" }
  // Last resort, labelled as such: the runner's clock (or a time counted from
  // the recording's start). It includes setup and late completion checks, so
  // it is an upper bound of the prompt-to-end time.
  if (runnerMs !== null) return { ms: capped(runnerMs), source: "runner-clock-upper-bound" }
  return null
}

// How each duration source is shown on the website inspector and the release viewer.
export const DURATION_SOURCE_LABELS = {
  "agent-time": "agent time",
  "claude-dashboard-session": "Claude dashboard session",
  "product-telemetry": "product telemetry",
  "recording-prompt-to-end": "recording, prompt to end",
  "recording-prompt-to-end-approximate": "recording, prompt to end (approximate)",
  "time-limit": "time limit (timed out)",
  "runner-clock-upper-bound": "runner clock, upper bound",
}

export function preferredBenchmarkDurationMs(record) {
  const totalDuration = finiteDuration(record?.totalElapsedMs)
    ?? finiteDuration(record?.durationMs)
  const agentDuration = finiteDuration(record?.agentElapsedMs)
    ?? finiteDuration(record?.agentDurationMs)
  if (agentDuration !== null) {
    // The runner stops a timed-out run at its limit, but the backend agent can
    // keep counting after the stop. The runner's time is what the user waited.
    if (record?.timedOut === true && totalDuration !== null && totalDuration < agentDuration) return totalDuration
    return agentDuration
  }

  const benchmarkDuration = finiteDuration(record?.benchmarkElapsedMs)
  if (benchmarkDuration !== null) return benchmarkDuration

  if (totalDuration !== null) return totalDuration

  const durationSeconds = finiteDuration(record?.durationSeconds)
  return durationSeconds === null ? null : durationSeconds * 1000
}
