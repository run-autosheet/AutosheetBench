"use client"

import { useLayoutEffect, useRef, useState } from "react"
import cn from "classnames"
import { DataTable } from "./chart-data"
import { chartOrder, scoreDurationData } from "./chart-rows"
import ChartTools, { ChartBrand } from "./chart-tools"
import {
  ConfigurationSlots,
  PAIR_MARK,
  PlatformBadgeMark,
  ProductLogoMark,
  ProductPairMark,
} from "./product-logo"
import { shownReasoning } from "./model-labels"
import {
  TASK_COUNT,
  PROFILES,
  TARGET_DURATION,
  TARGET_SCORE,
  configurationName,
  durationText,
  pickConfigurations,
  scoreText,
} from "./benchmark-model"
import {
  durationAxis,
  labelPoints,
  plotArea,
  scatterProfiles,
  scoreAxis,
} from "./scatter-layout"
import PlatformPills, { ConfigurationPills } from "./platform-pills"
import styles from "./benchmark.module.css"

// Every configuration on both platforms, whatever the toggles show, so the
// page's HTML holds all the results.

// From every configuration, so the time scale stays put when the toggles
// pick fewer.


// Named as the charts name it, so Excel's and Sheets' points read apart.
const accessibleLabel = (profile) => {
  const name = [
    configurationName(profile),
    profile.modelName,
    shownReasoning(profile),
  ]
    .filter(Boolean)
    .join(", ")
  return `${name}: ${scoreText(profile.score)}, ${durationText(profile.duration)}`
}
// The nominal mark size, and each marker's half extent for the layout: a
// round icon reserves 0.72 of it around its center, a pair its pill.
const MARK = 22
const extent = (profile, pad) =>
  profile.connector
    ? {
        halfX: (MARK * PAIR_MARK.width) / 2 + pad,
        halfY: (MARK * PAIR_MARK.height) / 2 + pad,
      }
    : { halfX: MARK * 0.72 + pad, halfY: MARK * 0.72 + pad }
const NAVIGATION_KEYS = [
  "ArrowLeft",
  "ArrowRight",
  "ArrowUp",
  "ArrowDown",
  "Home",
  "End",
]

// How far a platform badge reaches outside its marker, and the halo.
const BADGE_GAP = 3
const HALO = 2

// The configurations the toggles show, `profiles` being those of the
// platform and product type. With both platforms shown, each marker takes
// its platform's icon as a badge, so Excel and Sheets points read apart.
export default function ScoreDurationScatter({
  profiles: visible,
  platform,
  onPlatformChange,
  configurations,
  onConfigurationsChange,
  id,
  taskCount = TASK_COUNT,
}) {
  const DATA = scoreDurationData(chartOrder(visible))
  const TIME = durationAxis(visible)
  const profiles = pickConfigurations(visible, configurations)
  const [selectedId, setSelectedId] = useState(null)
  // The point Tab lands on: the last one focused, or the first.
  const [activeId, setActiveId] = useState(null)
  // A phone folds the key behind a button; each point keeps its logo.
  const [keyOpen, setKeyOpen] = useState(false)
  const pointRefs = useRef([])
  const panel = useRef(null)
  // Drawn at the chart's own width, one unit a pixel, so its axis gutter
  // and tick text match the overall score's at any width. The server
  // renders the desktop width; the first layout corrects it.
  const visual = useRef(null)
  const [width, setWidth] = useState(1296)
  useLayoutEffect(() => {
    const observer = new ResizeObserver(([entry]) =>
      setWidth(Math.round(entry.contentRect.width)),
    )
    observer.observe(visual.current)
    return () => observer.disconnect()
  }, [])
  const plot = plotArea(width, 430)
  const points = scatterProfiles(profiles)
  // The key holds every kind of point on the chart once, drawn as its points
  // are and named as the charts name it, so a point reads as its product
  // and its label as the model, in the charts' order.
  const products = chartOrder(points).filter(
    (profile, index, ordered) =>
      ordered.findIndex((other) => other.name === profile.name) === index,
  )
  const score = scoreAxis(points)
  const badged = platform === "all"
  const pad = badged ? BADGE_GAP : 0
  const x = (duration) =>
    plot.left + (duration / TIME.max) * (plot.right - plot.left)
  const y = (value) =>
    plot.bottom -
    ((value - score.min) / (100 - score.min)) * (plot.bottom - plot.top)
  // Every icon sits at its exact value. Labels go to the most cramped icons
  // first, the highest score breaking ties; icons are drawn lowest score
  // first, so where two overlap the better one stays whole.
  const bubbles = labelPoints(
    [...points]
      .sort((a, b) => b.score - a.score)
      .map((profile) => ({
        profile,
        label: profile.modelName,
        ...extent(profile, pad),
        bubbleX: x(profile.duration),
        bubbleY: y(profile.score),
      })),
    plot,
  ).reverse()
  const selected = bubbles.find((bubble) => bubble.profile.id === selectedId)
  const tabStop = bubbles.some((bubble) => bubble.profile.id === activeId)
    ? activeId
    : bubbles[0]?.profile.id
  function navigate(event, index) {
    if (event.key === "Escape") {
      setSelectedId(null)
      return
    }
    if (["Enter", " "].includes(event.key)) {
      event.preventDefault()
      setSelectedId(bubbles[index].profile.id)
      return
    }
    if (!NAVIGATION_KEYS.includes(event.key)) return
    event.preventDefault()
    const next =
      event.key === "Home"
        ? 0
        : event.key === "End"
          ? bubbles.length - 1
          : (index +
              (["ArrowLeft", "ArrowUp"].includes(event.key) ? -1 : 1) +
              bubbles.length) %
            bubbles.length
    pointRefs.current[next]?.focus()
  }

  // A card-colored halo just outside the marker's outline (a pill, the
  // brand's rounded square or a disk), which keeps overlapping icons apart.
  const halo = ({ profile, bubbleX, bubbleY }) => {
    if (profile.connector) {
      const width = MARK * PAIR_MARK.width + HALO * 2
      const height = MARK * PAIR_MARK.height + HALO * 2
      return (
        <rect
          className={styles.scatterHalo}
          x={bubbleX - width / 2}
          y={bubbleY - height / 2}
          width={width}
          height={height}
          rx={height / 2}
        />
      )
    }
    if (profile.branded) {
      const side = MARK + HALO * 2
      return (
        <rect
          className={styles.scatterHalo}
          x={bubbleX - side / 2}
          y={bubbleY - side / 2}
          width={side}
          height={side}
          rx={side * 0.24}
        />
      )
    }
    return (
      <circle
        className={styles.scatterHalo}
        cx={bubbleX}
        cy={bubbleY}
        r={MARK * 0.64 + HALO}
      />
    )
  }

  // The badge sits on the marker's lower right corner.
  const badge = ({ profile, bubbleX, bubbleY, halfX, halfY }) => (
    <PlatformBadgeMark
      platform={profile.platform}
      x={bubbleX + halfX - BADGE_GAP - (profile.connector ? 6 : 3)}
      y={bubbleY + halfY - BADGE_GAP - 3}
      size={10}
    />
  )

  // A product's mark centered on its bubble: its logo, or an assistant and
  // its MCP on a pill, with its platform's badge when both platforms show.
  const mark = (bubble) => (
    <>
      {bubble.profile.connector ? (
        <ProductPairMark
          productKey={bubble.profile.productKey}
          connector={bubble.profile.connector}
          branded={bubble.profile.branded}
          x={bubble.bubbleX}
          y={bubble.bubbleY}
          size={MARK}
        />
      ) : (
        <ProductLogoMark
          productKey={bubble.profile.productKey}
          x={bubble.bubbleX}
          y={bubble.bubbleY}
          size={MARK}
        />
      )}
      {badged && badge(bubble)}
    </>
  )

  // A point: its mark on its halo, its model's label, and a line to the
  // label where the label sits apart.
  const marker = (bubble) => (
    <>
      {bubble.labelLink && (
        <line
          className={styles.scatterLabelLink}
          x1={bubble.bubbleX}
          y1={bubble.bubbleY}
          x2={bubble.labelLink.x}
          y2={bubble.labelLink.y}
        />
      )}
      {halo(bubble)}
      {mark(bubble)}
      <text
        x={bubble.labelX}
        y={bubble.labelY}
        className={styles.scatterModelLabel}
      >
        {bubble.profile.modelName}
      </text>
    </>
  )

  // A product's mark for the key, at the chart's size.
  const keyMark = (profile) => {
    const { halfX, halfY } = extent(profile, pad)
    return (
      <svg
        viewBox={`0 0 ${halfX * 2} ${halfY * 2}`}
        width={halfX * 2}
        height={halfY * 2}
        aria-hidden="true"
        focusable="false"
      >
        {mark({ profile, bubbleX: halfX, bubbleY: halfY, halfX, halfY })}
      </svg>
    )
  }

  return (
    <div id={id} ref={panel} className={styles.leaderboardPanel}>
      <div className={styles.scatterIntro}>
        <div>
          <h3>Score vs duration</h3>
        </div>
        <div className={styles.scatterKey}>
          <div className={styles.chartControls} data-image="hide">
            <PlatformPills platform={platform} onChange={onPlatformChange} />
            <ConfigurationPills
              configurations={configurations}
              onChange={onConfigurationsChange}
            />
            <ChartTools
              anchor={id}
              file="score-vs-duration"
              image={panel}
              data={() => scoreDurationData(chartOrder(profiles))}
            />
          </div>
        </div>
        <button
          type="button"
          className={styles.keyToggle}
          aria-expanded={keyOpen}
          onClick={() => setKeyOpen(!keyOpen)}
          data-image="hide"
        >
          {keyOpen ? "Hide the legend" : "Show the legend"}
        </button>
        {/* The key reads with the chart, left under the title; the toggles
            keep to the right. */}
        <div
          className={cn(styles.scatterKeyItems, !keyOpen && styles.isFolded)}
        >
          <span>
            <i className={styles.quadrantKey} aria-hidden="true" />
            <span>Most attractive quadrant</span>
          </span>
          {products.map((profile) => (
            <span key={configurationName(profile)}>
              {keyMark(profile)}
              <span>{configurationName(profile)}</span>
            </span>
          ))}
          {/* The badges need no key of their own: each product's name
              gives its platform. */}
          <ChartBrand className={styles.keyBrand} />
        </div>
      </div>
      <div
        className={styles.scatterScroll}
        data-image="scroll"
        role="region"
        aria-label="Score versus duration chart"
        tabIndex={0}
      >
        <div
          ref={visual}
          className={styles.scatterVisual}
          onMouseLeave={() => setSelectedId(null)}
        >
          <svg
            className={styles.scatterPlot}
            viewBox={`0 0 ${plot.width} ${plot.height}`}
            // Its drawn height whatever the width, so the server's desktop
            // width doesn't leave a shorter chart that grows once measured
            // and moves everything under it, and links into the page.
            height={plot.height}
            role="group"
            aria-label={`Overall score versus average duration. Score axis starts at ${score.min} percent. Select an icon; use arrow keys to move between configurations.`}
            onClick={(event) => {
              if (!event.target.closest('[role="button"]')) setSelectedId(null)
            }}
          >
            {/* Read upward, beside the ticks, centered on the plot. */}
            <text
              x={-(plot.top + plot.bottom) / 2}
              y="34"
              transform="rotate(-90)"
              textAnchor="middle"
              className={styles.scatterAxisLabel}
            >
              Overall score →
            </text>
            {/* Under the grid and the points: 90% or more, under a
                minute, both edges on grid lines. */}
            <rect
              className={styles.scatterQuadrant}
              x={plot.left}
              y={plot.top}
              width={x(TARGET_DURATION) - plot.left}
              height={y(TARGET_SCORE) - plot.top}
            />
            {score.ticks.map((tick) => (
              <g key={tick}>
                {/* The lowest line is the baseline, as the overall
                    score's. */}
                <line
                  x1={plot.left}
                  x2={plot.right}
                  y1={y(tick)}
                  y2={y(tick)}
                  className={cn(
                    styles.scatterGrid,
                    tick === score.min && styles.scatterBaseline,
                  )}
                />
                <text
                  x={plot.left - 8}
                  y={y(tick)}
                  textAnchor="end"
                  dominantBaseline="central"
                  className={styles.scatterTick}
                >
                  {`${tick}%`}
                </text>
              </g>
            ))}
            {TIME.ticks.map((tick) => (
              <g key={tick}>
                <line
                  x1={x(tick)}
                  x2={x(tick)}
                  y1={plot.top}
                  y2={plot.bottom}
                  className={styles.scatterGrid}
                />
                <text
                  x={x(tick)}
                  y={plot.bottom + 25}
                  textAnchor="middle"
                  className={styles.scatterTick}
                >
                  {durationText(tick)}
                </text>
              </g>
            ))}
            {/* Centered under the plot, as the score title beside it. */}
            <text
              x={(plot.left + plot.right) / 2}
              y={plot.bottom + 54}
              textAnchor="middle"
              className={styles.scatterAxisLabel}
            >
              Average duration per task →
            </text>
            {/* The quadrant's two inner edges, over the grid, so it reads
                as a zone and not only a tint. */}
            <path
              className={styles.scatterQuadrantEdge}
              d={`M${x(TARGET_DURATION)} ${plot.top}V${y(TARGET_SCORE)}H${plot.left}`}
            />
            {selected && (
              <g className={styles.scatterGuide} pointerEvents="none">
                <line
                  x1={plot.left}
                  x2={selected.bubbleX}
                  y1={selected.bubbleY}
                  y2={selected.bubbleY}
                />
                <line
                  x1={selected.bubbleX}
                  x2={selected.bubbleX}
                  y1={selected.bubbleY}
                  y2={plot.bottom}
                />
              </g>
            )}
            {bubbles.map((bubble, index) => {
              const { profile, bubbleX, bubbleY, halfX, halfY } = bubble
              return (
                <g
                  key={profile.id}
                  role="button"
                  tabIndex={profile.id === tabStop ? 0 : -1}
                  ref={(element) => {
                    pointRefs.current[index] = element
                  }}
                  className={cn(
                    styles.scatterPoint,
                    profile.id === selectedId && styles.isSelected,
                    bubble.labelHidden && styles.labelOnHover,
                  )}
                  aria-label={accessibleLabel(profile)}
                  onMouseEnter={() => setSelectedId(profile.id)}
                  onFocus={() => {
                    setActiveId(profile.id)
                    setSelectedId(profile.id)
                  }}
                  onBlur={() => setSelectedId(null)}
                  onClick={() => setSelectedId(profile.id)}
                  onKeyDown={(event) => navigate(event, index)}
                >
                  <rect
                    x={bubbleX - halfX}
                    y={bubbleY - halfY}
                    width={halfX * 2}
                    height={halfY * 2}
                    fill="transparent"
                  />
                  {marker(bubble)}
                </g>
              )
            })}
            {/* The hovered or focused icon, drawn again over its neighbours. */}
            {selected && (
              <g
                className={cn(styles.scatterPoint, styles.isSelected)}
                pointerEvents="none"
                aria-hidden="true"
              >
                {marker(selected)}
              </g>
            )}
          </svg>
          {selected && (
            <div
              className={styles.scatterTooltip}
              role="status"
              data-image="hide"
              style={{
                left: `clamp(8px, ${((selected.bubbleX + selected.halfX + 9) / plot.width) * 100}%, calc(100% - 288px))`,
                top: `${((selected.bubbleY > plot.height / 2 ? selected.bubbleY - 116 : selected.bubbleY + 24) / plot.height) * 100}%`,
              }}
            >
              <ConfigurationSlots profile={selected.profile} />
              <div className={styles.scatterTooltipMetrics}>
                <span>
                  <strong>{scoreText(selected.profile.score)}</strong> overall
                </span>
                <span>
                  <strong>{durationText(selected.profile.duration)}</strong>{" "}
                  average
                </span>
              </div>
              <small>{`${selected.profile.scored}/${taskCount} tasks scored`}</small>
            </div>
          )}
        </div>
      </div>
      <DataTable
        caption="Score vs duration: every configuration, on both platforms"
        data={DATA}
      />
    </div>
  )
}
