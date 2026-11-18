"use client"

import { useRef, useState } from "react"
import cn from "classnames"
import CapabilityRanks from "./capability-ranks"
import { DataTable } from "./chart-data"
import { capabilityData, chartOrder, rankData } from "./chart-rows"
import ChartTools, { ChartBrand } from "./chart-tools"
import { ArrowDown, ArrowUp, ArrowUpDown } from "./icons"
import PlatformPills, { ConfigurationPills } from "./platform-pills"
import { Pills } from "./pills"
import { ProductLogo, seriesStyles } from "./product-logo"
import { RADAR_VIEWBOX, RadarAxes, RadarSeries, nearestSeries } from "./radar"
import {
  TASK_COUNT,
  CATEGORY_DETAILS,
  CATEGORY_TOTALS,
  PROFILES,
  TARGET_SCORE,
  configurationName,
  durationText,
  finite,
  groupByKind,
  pickConfigurations,
  scoreText,
  sortProfiles,
  shownReasoning,
} from "./benchmark-model"
import styles from "./benchmark.module.css"

const COLUMNS = [
  { key: "score", label: "Overall", title: "Overall task success (%)" },
  {
    key: "duration",
    label: "Duration",
    title: "Average task duration (lower is better)",
  },
  ...CATEGORY_DETAILS.map(({ name }, index) => ({
    key: index,
    label: name,
    title: `${name} task success (%)`,
  })),
]

// A diverging scale around 90%: green deepens to 100%, coral deepens down
// to 30%, and a score near 90% stays plain.
const MIDPOINT = TARGET_SCORE
const tint = (value) => {
  const high = value >= MIDPOINT
  const depth = high
    ? (value - MIDPOINT) / (100 - MIDPOINT)
    : Math.min(1, (MIDPOINT - value) / (MIDPOINT - 30))
  return {
    "--tint": `${Math.round(depth * 100)}%`,
    "--tint-color": high
      ? "var(--benchmark-score-high)"
      : "var(--benchmark-score-low)",
  }
}

// Every configuration on both platforms, whatever the toggles show, so the
// page's HTML holds all the results.


const VIEWS = [
  { value: "radar", label: "Radar" },
  { value: "table", label: "Table" },
  { value: "ranks", label: "Ranks" },
]

// Up to this many shapes are filled. Past it, their fills stack into a dark
// middle that hides the rings, so they are drawn as outlines and only the
// shape brought forward is filled.
const MAX_FILLED = 4

// The same configurations as the table, each a shape over the eight
// categories, beside their scores: a row each, in the table's groups and
// order, named with the shape's line as the radar's key, and a column per
// category in the radar's clockwise order. Hovering or focusing a row's
// name, or pointing inside the radar's rings, brings that configuration's
// shape and row forward and fades the others; selecting either keeps it
// so.
function CapabilityRadar({ groups, headed, taskCount, categoryTotals }) {
  const [pinned, setPinned] = useState(null)
  const [hovered, setHovered] = useState(null)
  const profiles = groups.flatMap((group) => group.items)
  const styled = seriesStyles(profiles)
  const series = profiles.map((profile, index) => ({
    profile,
    ...styled[index],
  }))
  // A pinned configuration the toggles no longer show focuses nothing, or
  // every shown one would be faded.
  const shown = (id) => profiles.some((profile) => profile.id === id)
  const focused = shown(hovered) ? hovered : shown(pinned) ? pinned : null
  // The focused shape draws last, on top of the others.
  const drawn = [
    ...series.filter(({ profile }) => profile.id !== focused),
    ...series.filter(({ profile }) => profile.id === focused),
  ]
  // The configuration whose shape the pointer is on, from the pointer's
  // place in the radar's own coordinates.
  const seriesAt = (event) => {
    const svg = event.currentTarget
    const at = new DOMPoint(event.clientX, event.clientY).matrixTransform(
      svg.getScreenCTM().inverse(),
    )
    const index = nearestSeries(
      profiles.map((profile) => profile.categories),
      [at.x, at.y],
    )
    return index === null ? null : profiles[index].id
  }
  const swatch = ({ color, dash }) => (
    <svg width="22" height="12" aria-hidden="true" focusable="false">
      <line
        x1="1"
        y1="6"
        x2="21"
        y2="6"
        stroke={color}
        strokeWidth="2.5"
        strokeDasharray={dash}
      />
    </svg>
  )
  return (
    <div className={styles.capabilityRadar}>
      <svg
        className={styles.radar}
        viewBox={RADAR_VIEWBOX}
        onMouseMove={(event) => setHovered(seriesAt(event))}
        onMouseLeave={() => setHovered(null)}
        data-outlined={profiles.length > MAX_FILLED ? "" : undefined}
        onClick={(event) => {
          const id = seriesAt(event)
          if (id) setPinned((current) => (current === id ? null : id))
        }}
        role="img"
        aria-label="Capability scores of the configurations shown, one shape each. The table beside it gives the scores."
      >
        <RadarAxes />
        {drawn.map(({ profile, color, dash }) => (
          <RadarSeries
            key={profile.id}
            values={profile.categories}
            name={`${configurationName(profile)}, ${profile.modelName}`}
            color={color}
            dash={dash}
            className={cn(
              focused && profile.id !== focused && styles.radarDimmed,
              profile.id === focused && styles.radarFocused,
            )}
          />
        ))}
      </svg>
      <div
        className={styles.radarScores}
        data-image="scroll"
        onMouseLeave={() => setHovered(null)}
      >
        <table data-focused={focused ? "" : undefined}>
          <caption className={styles.visuallyHidden}>
            Overall and capability scores of the configurations shown, a row
            each, the capabilities in the radar&apos;s order. Select a
            configuration to keep its shape forward.
          </caption>
          <thead>
            <tr>
              <td />
              {["Overall", ...CATEGORY_DETAILS.map(({ name }) => name)].map(
                (name) => (
                  <th key={name} scope="col" title={name}>
                    <span>{name}</span>
                  </th>
                ),
              )}
            </tr>
          </thead>
          {groups.map((group) => (
            <tbody key={group.key}>
              {headed && (
                <tr className={styles.radarScoresGroups}>
                  <th scope="rowgroup" colSpan={CATEGORY_DETAILS.length + 2}>
                    {group.label}
                  </th>
                </tr>
              )}
              {group.items.map((profile) => {
                const style = series.find((item) => item.profile === profile)
                return (
                  <tr
                    key={profile.id}
                    className={cn(
                      styles.integrationResult,
                      profile.id === focused && styles.radarScoresFocused,
                    )}
                  >
                    <th scope="row">
                      <button
                        type="button"
                        aria-pressed={pinned === profile.id}
                        onClick={() =>
                          setPinned((current) =>
                            current === profile.id ? null : profile.id,
                          )
                        }
                        onMouseEnter={() => setHovered(profile.id)}
                        onFocus={() => setHovered(profile.id)}
                        onBlur={() => setHovered(null)}
                      >
                        {swatch(style)}
                        <ProductLogo profile={profile} size={20} />
                        <span>
                          <strong>{configurationName(profile)}</strong>
                          <span>
                            {profile.modelName}
                            {shownReasoning(profile) && (
                              <em> ({shownReasoning(profile)})</em>
                            )}
                          </span>
                        </span>
                      </button>
                    </th>
                    <MetricCell
                      value={profile.score}
                      overall
                      title={`${profile.scored} of ${taskCount} tasks scored`}
                    />
                    {CATEGORY_DETAILS.map(({ name }, i) => (
                      <MetricCell
                        key={name}
                        value={profile.categories[i]}
                        title={`${name}: ${scoreText(profile.categories[i])}, ${profile.categoryScored[i]} of ${categoryTotals[i]} tasks scored`}
                      />
                    ))}
                  </tr>
                )
              })}
            </tbody>
          ))}
        </table>
      </div>
    </div>
  )
}

function MetricCell({ value, duration, title, overall }) {
  const tinted = !duration && finite(value)
  return (
    <td
      className={cn(
        styles.metricCell,
        overall && styles.overallCell,
        duration && styles.durationCell,
        tinted && styles.tintedCell,
      )}
      style={tinted ? tint(value) : undefined}
      title={title}
    >
      {finite(value)
        ? duration
          ? durationText(value)
          : value.toFixed(1)
        : "—"}
    </td>
  )
}

// Every configuration's overall, category and duration results, grouped as
// the score chart groups its columns: by platform, then kind of product.
export default function Leaderboard({
  profiles: visible,
  platform,
  onPlatformChange,
  configurations,
  onConfigurationsChange,
  id,
  taskCount = TASK_COUNT,
  categoryTotals = CATEGORY_TOTALS,
}) {
  const [sort, setSort] = useState({ key: "score", direction: "desc" })
  const [view, setView] = useState("radar")
  const DATA = capabilityData(chartOrder(visible))
  const panel = useRef(null)
  // Once the scores slide under the sticky names, the names column takes an
  // edge.
  const [scrolled, setScrolled] = useState(false)
  const shown = sortProfiles(pickConfigurations(visible, configurations), sort)
  // Sorting happens inside each group, never across them.
  const groups = groupByKind(shown, platform)
  // One group, as Excel's sidebar add-ons, needs no heading: the toggles
  // already name it.
  const headed = groups.length > 1
  // The configurations shown, in the table's order.
  const items = groups.flatMap((group) => group.items)
  // A row's rank in its group by the sorted column, as its cell reads, so
  // configurations that read the same share one ("1, 1, 3").
  const shownValue = (profile) => {
    const value =
      sort.key === "score" || sort.key === "duration"
        ? profile[sort.key]
        : profile.categories[sort.key]
    if (!finite(value)) return null
    return sort.key === "duration" ? durationText(value) : value.toFixed(1)
  }
  const rankIn = (group, profile) =>
    group.items.findIndex(
      (other) => shownValue(other) === shownValue(profile),
    ) + 1
  const changeSort = (key) =>
    setSort((current) => ({
      key,
      direction:
        current.key === key
          ? current.direction === "desc"
            ? "asc"
            : "desc"
          : key === "duration"
            ? "asc"
            : "desc",
    }))

  return (
    <div id={id} ref={panel} className={styles.leaderboardPanel}>
      <div className={styles.leaderboardCaption}>
        <div>
          <h3>Score by capability</h3>
          {/* The tint key reads on the tables, not the ranks. */}
          {view !== "ranks" && (
            <p className={styles.scoreKey}>
              <span>{`Below ${MIDPOINT}%`}</span>
              <span className={styles.scoreRamp} aria-hidden="true" />
              <span>{`Above ${MIDPOINT}%`}</span>
            </p>
          )}
          {view === "ranks" && (
            <p className={styles.scoreKey}>
              Rank among the configurations shown; ties share a rank.
            </p>
          )}
        </div>
        <div className={styles.chartAside}>
          <div className={styles.chartControls} data-image="hide">
            <Pills
              choices={VIEWS}
              value={view}
              onChange={setView}
              label="View"
            />
            <PlatformPills platform={platform} onChange={onPlatformChange} />
            <ConfigurationPills
              configurations={configurations}
              onChange={onConfigurationsChange}
            />
            {/* The rows in the table's order, as the radar's key lists
                them; the ranks view saves its ranks. */}
            <ChartTools
              anchor={id}
              file={
                view === "ranks" ? "rank-by-capability" : "score-by-capability"
              }
              image={panel}
              data={() =>
                view === "ranks" ? rankData(items) : capabilityData(items)
              }
            />
          </div>
          <ChartBrand />
        </div>
      </div>
      {view === "radar" ? (
        <CapabilityRadar groups={groups} headed={headed} taskCount={taskCount} categoryTotals={categoryTotals} />
      ) : view === "ranks" ? (
        <CapabilityRanks profiles={items} />
      ) : (
        <div
          className={styles.tableScroll}
          data-image="scroll"
          role="region"
          aria-label="Benchmark leaderboard. Scroll to explore all configurations and columns."
          tabIndex={0}
          data-scrolled={scrolled || undefined}
          onScroll={(event) => setScrolled(event.currentTarget.scrollLeft > 0)}
        >
          <table className={styles.integrationTable}>
            <caption className={styles.visuallyHidden}>
              Overall score, average task duration, and eight capability scores,
              grouped by kind of product. Select a column heading to sort within
              each group.
            </caption>
            <colgroup>
              <col className={styles.methodCol} />
              <col className={styles.overallCol} />
              <col className={styles.durationCol} />
              {CATEGORY_DETAILS.map(({ name }) => (
                <col className={styles.categoryCol} key={name} />
              ))}
            </colgroup>
            <thead>
              <tr>
                <th scope="col">Product / model</th>
                {COLUMNS.map((column) => (
                  <th
                    key={column.key}
                    scope="col"
                    className={cn(
                      sort.key === column.key && styles.sortedColumn,
                    )}
                    aria-sort={
                      sort.key === column.key
                        ? sort.direction === "desc"
                          ? "descending"
                          : "ascending"
                        : "none"
                    }
                  >
                    <button
                      type="button"
                      onClick={() => changeSort(column.key)}
                      title={column.title}
                      aria-label={`Sort by ${column.title}`}
                    >
                      {/* The arrow hangs off the label, so the label centers
                        over its column's tiles. */}
                      <span>
                        {column.label}
                        {sort.key === column.key ? (
                          sort.direction === "desc" ? (
                            <ArrowDown size={11} />
                          ) : (
                            <ArrowUp size={11} />
                          )
                        ) : (
                          <ArrowUpDown
                            className={styles.inactiveSort}
                            size={10}
                          />
                        )}
                      </span>
                    </button>
                  </th>
                ))}
              </tr>
            </thead>
            {groups.map((group) => (
              <tbody key={group.key}>
                {headed && (
                  <tr className={styles.integrationGroup}>
                    <th scope="rowgroup" colSpan={COLUMNS.length + 1}>
                      <span>{group.label}</span>
                    </th>
                  </tr>
                )}
                {group.items.map((profile) => (
                  <tr key={profile.id} className={styles.integrationResult}>
                    <th scope="row" className={styles.methodCell}>
                      <div className={styles.methodIdentity}>
                        <span className={styles.methodRank}>
                          {rankIn(group, profile)}
                        </span>
                        <ProductLogo profile={profile} />
                        <span className={styles.methodCopy}>
                          <strong>{configurationName(profile)}</strong>
                          <span>
                            {profile.modelName}
                            {shownReasoning(profile) && (
                              <em> ({shownReasoning(profile)})</em>
                            )}
                          </span>
                          {/* Coverage shows only when a configuration missed
                            tasks. */}
                          {profile.scored < taskCount && (
                            <small>{`${profile.scored}/${taskCount} scored`}</small>
                          )}
                        </span>
                      </div>
                    </th>
                    <MetricCell
                      value={profile.score}
                      overall
                      title={`${profile.scored} of ${taskCount} tasks scored`}
                    />
                    <MetricCell
                      value={profile.duration}
                      duration
                      title={`Average over ${profile.timedRuns} timed runs, each from prompt submission to final answer.`}
                    />
                    {CATEGORY_DETAILS.map(({ name }, i) => (
                      <MetricCell
                        key={name}
                        value={profile.categories[i]}
                        title={`${name}: ${profile.categoryScored[i]} of ${categoryTotals[i]} tasks scored`}
                      />
                    ))}
                  </tr>
                ))}
              </tbody>
            ))}
          </table>
        </div>
      )}
      <DataTable
        caption="Score by capability: every configuration, on both platforms"
        data={DATA}
      />
      {view === "ranks" && (
        <DataTable
          caption="Rank by capability: the configurations shown"
          data={rankData(items)}
        />
      )}
    </div>
  )
}
