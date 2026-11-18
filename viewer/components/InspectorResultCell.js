"use client"

// Matrix presentation shared by the internal inspector and the static public view.
// No registry loading, evaluation requests or historical writes occur here.
import { Fragment } from "react"
import { Video } from "lucide-react"
import { DURATION_SOURCE_LABELS, preferredBenchmarkDurationMs } from "../../lib/duration.mjs"

const styles = {
  page: { minHeight: "100vh", padding: "22px 24px", background: "var(--eb-bg)", color: "var(--eb-text)" },
  muted: { color: "var(--eb-text2)" },
  panel: { border: "1px solid var(--eb-border)", borderRadius: 8, background: "var(--eb-bg2)" },
  button: { border: "1px solid var(--eb-border)", borderRadius: 6, background: "var(--eb-bg3)", color: "var(--eb-text)", padding: "6px 10px", cursor: "pointer", fontSize: 12 },
  primary: { border: "1px solid #1f6feb", borderRadius: 6, background: "#1f6feb", color: "white", padding: "6px 10px", cursor: "pointer", fontSize: 12 },
}

function formatScore(value) {
  return typeof value === "number" && Number.isFinite(value) ? `${value % 1 === 0 ? value.toFixed(0) : value.toFixed(1)}%` : "—"
}

function formatDuration(value) {
  return typeof value === "number" && Number.isFinite(value) ? `${(value / 1000).toFixed(1)}s` : "—"
}

function statusForCell(cell) {
  if (!cell) return { label: "MISSING", color: "var(--eb-text2)", background: "var(--eb-bg3)" }
  if (cell.passed === true) return { label: "PASS", color: "var(--eb-green)", background: "rgba(46, 160, 67, .20)" }
  if (cell.passed === false) return { label: "FAIL", color: "var(--eb-red)", background: "rgba(248, 81, 73, .20)" }
  if (!cell.outputAvailable) return { label: "MISSING", color: "var(--eb-red)", background: "rgba(248, 81, 73, .12)" }
  if (cell.timedOut) return { label: "TIMEOUT", color: "var(--eb-orange)", background: "rgba(240, 136, 62, .16)" }
  if (cell.evaluationStatus === 'inconclusive') return { label: 'INCONCLUSIVE', color: 'var(--eb-yellow)', background: 'rgba(210, 153, 34, .16)' }
  if (cell.incomplete) return { label: "INCOMPLETE", color: "var(--eb-yellow)", background: "rgba(210, 153, 34, .16)" }
  if (cell.evaluationStatus === "evaluated") return { label: "UNKNOWN", color: "var(--eb-yellow)", background: "rgba(210, 153, 34, .12)" }
  return { label: "RECORDED", color: "var(--eb-blue)", background: "rgba(88, 166, 255, .10)" }
}

function evidenceLabel(value) {
  return {
    outputWorkbook: "output workbook",
    transcript: "transcript / trace",
    actualDocument: "actual screenshot",
    actualChartImage: "actual chart image",
    expectedDocument: "expected screenshot",
    expectedChartImage: "expected chart image",
    qualityJudgment: "quality judgment",
  }[value] || String(value || "unknown evidence")
}

function missingEvidenceText(cell) {
  return (Array.isArray(cell?.missingEvidence) ? cell.missingEvidence : []).map(evidenceLabel)
}

function Cell({ cell, onDetail, onRanking, onEvaluate, evaluating, onClear, clearing, onManualOverride, overriding, showToolUsage, showFullStats, publicSnapshot = false }) {
  if (!cell) {
    // Task with no recorded run for this configuration — render an empty
    // matrix slot so every task still appears as a row.
    return (
      <td style={{ borderTop: "1px solid var(--eb-border)", borderLeft: "1px solid var(--eb-border)", padding: "9px 10px", verticalAlign: "top", minWidth: 170, background: "var(--eb-bg3)" }}>
        <div style={{ ...styles.muted, fontSize: 16, fontWeight: 700 }}>—</div>
        <div style={{ ...styles.muted, fontSize: 10, marginTop: 3 }}>no run</div>
      </td>
    )
  }
  const status = statusForCell(cell)
  const hasScore = typeof cell.effectiveScore === "number" && Number.isFinite(cell.effectiveScore)
  const deterministicOnly = String(cell.evaluationMode || "").trim().toLowerCase().startsWith("deterministic") || cell.judgeType === "Deterministic"
  const qualityRequired = cell.qualityRequired === true
  const templateAnalysis = cell.templateAnalysis && typeof cell.templateAnalysis === "object"
    ? cell.templateAnalysis
    : null
  const benchmarkDurationMs = publicSnapshot ? cell.publicDurationMs : preferredBenchmarkDurationMs(cell)
  const recordingUrl = cell.recordingUrl || (cell.recordingAvailable
    ? `/api/benchmark-inspector/artifact?${new URLSearchParams({ kind: "actual", configId: cell.runConfigId, taskId: cell.taskId, artifact: "recording" }).toString()}`
    : null)
  return (
    <td
      role="button"
      tabIndex={0}
      onClick={onDetail}
      onKeyDown={(event) => { if (event.key === "Enter" || event.key === " ") { event.preventDefault(); onDetail() } }}
      title="Open recorded task details"
      className="transition-shadow hover:ring-1 hover:ring-inset hover:ring-cyan-400 focus-visible:ring-1 focus-visible:ring-inset focus-visible:ring-cyan-400"
      style={{ borderTop: "1px solid var(--eb-border)", borderLeft: "1px solid var(--eb-border)", padding: "9px 10px", verticalAlign: "top", minWidth: 190, cursor: "pointer", background: status.background, outline: "none" }}
    >
      <div style={{ display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 8 }}>
        <div style={{ display: "flex", alignItems: "baseline", gap: 8 }}>
          <span style={{ fontSize: 22, fontWeight: 800, letterSpacing: ".01em", color: hasScore ? status.color : "var(--eb-text2)" }}>{hasScore ? formatScore(cell.effectiveScore) : status.label}</span>
          {(hasScore || qualityRequired) && <span style={{ ...styles.muted, fontSize: 10 }}>{publicSnapshot ? "Saved score" : qualityRequired ? `Quality ${formatScore(cell.qualityScore)} · Det ${formatScore(cell.deterministicScore)}` : deterministicOnly ? `Det ${formatScore(cell.deterministicScore)}` : `LLM ${formatScore(cell.judgeScore)} · Det ${formatScore(cell.deterministicScore)}`}</span>}
          {cell.minorIssues > 0 && (
            <span
              title={`${cell.minorIssues} minor issue${cell.minorIssues === 1 ? "" : "s"}: one deducts 50 points; two or more score 0. Main failures always score 0.`}
              style={{ fontSize: 10, fontWeight: 700, color: "#fbbf24", border: "1px solid #b45309", borderRadius: 4, padding: "0 4px" }}
            >
              {cell.minorIssues} minor
            </span>
          )}
          {cell.reviewVerdict === "disagree" && (
            <span
              title={`Judge disagrees with the checks: ${cell.reviewReason || ""}`}
              style={{ fontSize: 10, fontWeight: 700, color: "#fbbf24", border: "1px solid #b45309", borderRadius: 4, padding: "0 4px" }}
            >
              judge disagrees
            </span>
          )}
        </div>
        {recordingUrl && (
          <a
            href={recordingUrl}
            target="_blank"
            rel="noreferrer"
            aria-label={`Open recording for ${cell.taskId}`}
            title="Open run recording"
            onClick={(event) => event.stopPropagation()}
            onKeyDown={(event) => event.stopPropagation()}
            className="inline-flex shrink-0 items-center justify-center rounded border border-white/10 bg-black/15 text-gray-400 hover:border-cyan-400/60 hover:text-cyan-300 focus-visible:outline-none focus-visible:ring-1 focus-visible:ring-cyan-400"
            style={{ width: 24, height: 24 }}
          >
            <Video size={14} strokeWidth={1.8} aria-hidden="true" />
          </a>
        )}
      </div>
      {cell.taskRevision?.status === 'outdated' && <div title={cell.taskRevision.message} style={{ color: '#fbbf24', fontSize: 10, marginTop: 4 }}>⚠ Task updated</div>}
      <div style={{ ...styles.muted, fontSize: 11, marginTop: 2 }}>{publicSnapshot && cell.publicRecordingOnly && benchmarkDurationMs == null ? "Duration unavailable" : formatDuration(benchmarkDurationMs)}{publicSnapshot ? ` · ${cell.publicAnswerText || "Final answer unavailable"}${cell.publicImagesExpected === false ? "" : ` · ${cell.publicImages || 0} images`}` : ` · ${cell.loops ?? "—"} loops`}</div>
      {publicSnapshot ? <div style={{ ...styles.muted, fontSize: 10, marginTop: 3 }}>{cell.publicRecordingOnly ? (benchmarkDurationMs == null ? "" : `${cell.publicDurationSource === "runner-clock-upper-bound" ? "≤ " : ""}${DURATION_SOURCE_LABELS[cell.publicDurationSource] || "recording"}${Number.isFinite(cell.publicDurationBoundSeconds) ? ` ±${cell.publicDurationBoundSeconds.toFixed(1)}s` : ""}`) : cell.publicTimingPolicy || "Timing not recorded"}</div> : <div style={{ ...styles.muted, fontSize: 10, marginTop: 2 }}>limit {cell.timeLimitSeconds != null ? `${cell.timeLimitSeconds}s` : "—"} · {cell.finishedInTime === false ? <span style={{ color: "var(--eb-orange)" }}>over time</span> : cell.finishedInTime === true ? "in time" : "time —"}{cell.agentElapsedMs != null && cell.totalElapsedMs != null ? ` · conversation ${formatDuration(cell.totalElapsedMs)}` : ""}</div>}
      {templateAnalysis && (
        <div style={{ color: templateAnalysis.failed ? "#fca5a5" : "var(--eb-text2)", fontSize: 10, marginTop: 3 }}>
          Templates: {templateAnalysis.detectedTemplateCount ?? "—"} cluster{templateAnalysis.detectedTemplateCount === 1 ? "" : "s"} · top {Math.min(templateAnalysis.detectedTemplateCount ?? templateAnalysis.maxTemplates ?? 0, templateAnalysis.maxTemplates ?? 0) || "—"} cover {Math.round((templateAnalysis.topTemplateCoverage || 0) * 1000) / 10}%
        </div>
      )}
      {cell.executionWarning && <div title={cell.executionWarning} style={{ color: "var(--eb-yellow)", fontSize: 10, marginTop: 3 }}>⚠ runner warning{cell.executionFailurePhase ? ` · ${cell.executionFailurePhase}` : ""}</div>}
      {missingEvidenceText(cell).length > 0 && <div style={{ color: "var(--eb-yellow)", fontSize: 10, marginTop: 3 }}>missing: {missingEvidenceText(cell).join(", ")}</div>}
      <div style={{ display: "flex", gap: 5, marginTop: 7 }}>
        <button type="button" onClick={(event) => { event.stopPropagation(); onDetail() }} style={styles.button}>Details</button>
      </div>
    </td>
  )
}

export { Cell }
