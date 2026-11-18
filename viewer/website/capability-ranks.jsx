"use client"

import { useLayoutEffect, useRef, useState } from "react"
import cn from "classnames"
import { ProductLogoMark, seriesStyles } from "./product-logo"
import { nameLines } from "./radar"
import {
  CATEGORY_DETAILS,
  configurationName,
  rankByCategory,
  scoreText,
  shownReasoning,
} from "./benchmark-model"
import styles from "./benchmark.module.css"

// A row per configuration, under the capabilities' names, between the
// labels; each label holds a logo, then the name and the model.
const ROW = 26
const HEAD = 52
const LOGO = 18
const label = (profile) => [configurationName(profile), profile.modelName]
// A label's width, a little over the semibold name's and the regular
// model's at the chart's 12px, which the server can't measure.
const labelWidth = (profile) =>
  configurationName(profile).length * 6.2 +
  ` · ${profile.modelName}`.length * 5.6
// The room a capability's name needs beside its neighbors': "Data
// operations" is 96px wide, "processing" beside it 68px.
const COLUMN = 90

// Each configuration's rank in each capability among those shown, as
// Artificial Analysis compares models across its capability indexes: a
// line through its place in every column, each place numbered with its
// rank, ties sharing one. The labels name the configurations in the first
// column's order on the left and, where the chart has room, the last's on
// the right. Hovering or
// focusing a line or a label brings it forward and fades the others;
// selecting it keeps it so. `profiles` come in the table's order, which
// sets their colors as in the radar view. Screen readers get the left
// labels alone: the card's data table lists every rank, which the lines'
// numbers would only read out of order.
export default function CapabilityRanks({ profiles }) {
  const [pinned, setPinned] = useState(null)
  const [hovered, setHovered] = useState(null)
  // Drawn at its own width, one unit a pixel, as the score vs duration
  // chart, so its text keeps its size; the server renders the desktop
  // width and the first layout corrects it.
  const box = useRef(null)
  const [width, setWidth] = useState(1248)
  useLayoutEffect(() => {
    const observer = new ResizeObserver(([entry]) =>
      setWidth(Math.round(entry.contentRect.width)),
    )
    observer.observe(box.current)
    return () => observer.disconnect()
  }, [])
  const ranks = rankByCategory(profiles)
  const styled = seriesStyles(profiles)
  const series = profiles.map((profile, index) => ({
    profile,
    ...styled[index],
    places: ranks.map((column) => column.get(profile.id)),
  }))
  // A pinned configuration the toggles no longer show focuses nothing, or
  // every shown one would be faded.
  const shown = (id) => profiles.some((profile) => profile.id === id)
  const focused = shown(hovered) ? hovered : shown(pinned) ? pinned : null
  // The focused line draws last, over the others.
  const drawn = [
    ...series.filter(({ profile }) => profile.id !== focused),
    ...series.filter(({ profile }) => profile.id === focused),
  ]
  // Each side's labels as wide as the longest, and the right side only
  // where both leave every capability its room.
  const labels = Math.ceil(Math.max(...profiles.map(labelWidth))) + LOGO + 20
  const both =
    (width - 2 * labels - 48) / (CATEGORY_DETAILS.length - 1) >= COLUMN
  const end = both ? width - labels : width
  const first = labels + 24
  const last = end - 24
  const x = (category) =>
    first + (category * (last - first)) / (CATEGORY_DETAILS.length - 1)
  const y = (position) => HEAD + ROW / 2 + position * ROW
  const height = HEAD + profiles.length * ROW + 8
  const dimmed = (profile) => focused && profile.id !== focused
  const toggle = (id) => setPinned((current) => (current === id ? null : id))
  // From the left label through each column to the right label, each step
  // a curve that leaves and meets its columns level.
  const path = (places) => {
    const points = places.map(({ position }, category) => [
      x(category),
      y(position),
    ])
    return [
      `M ${labels} ${points[0][1]}`,
      `L ${points[0][0]} ${points[0][1]}`,
      ...points.slice(1).map(([px, py], index) => {
        const [qx, qy] = points[index]
        const half = (px - qx) / 2
        return `C ${qx + half} ${qy} ${px - half} ${py} ${px} ${py}`
      }),
      `L ${both ? end : last} ${points.at(-1)[1]}`,
    ].join(" ")
  }
  // A configuration's label on either side, its logo against the lines.
  const sideLabel = ({ profile, places }, side) => {
    const [name, model] = label(profile)
    const row = y(places[side === "left" ? 0 : places.length - 1].position)
    const logoX =
      side === "left" ? labels - LOGO / 2 - 6 : width - labels + LOGO / 2 + 6
    const textX =
      side === "left" ? labels - LOGO - 14 : width - labels + LOGO + 14
    return (
      <>
        <ProductLogoMark
          productKey={profile.productKey}
          connector={profile.connector}
          x={logoX}
          y={row}
          size={LOGO}
        />
        <text
          x={textX}
          y={row}
          textAnchor={side === "left" ? "end" : "start"}
          dominantBaseline="central"
          className={styles.rankLabel}
        >
          <tspan className={styles.rankName}>{name}</tspan>
          {` · ${model}`}
        </text>
      </>
    )
  }
  return (
    <div className={styles.rankScroll} data-image="scroll">
      {/* At least as wide as the labels and a column's room for each
          capability: narrower, it scrolls. */}
      <div
        ref={box}
        className={styles.rankBox}
        style={{
          minWidth: labels + 48 + (CATEGORY_DETAILS.length - 1) * COLUMN + 48,
        }}
      >
        <svg
          className={styles.rankChart}
          viewBox={`0 0 ${width} ${height}`}
          role="group"
          aria-label="Rank of each configuration shown in each capability, listed in the data table that follows. Select a name to bring its line forward."
          onMouseLeave={() => setHovered(null)}
        >
          {CATEGORY_DETAILS.map(({ name }, category) => {
            const lines = nameLines(name)
            return (
              <text
                key={name}
                x={x(category)}
                y={HEAD - 14 - (lines.length - 1) * 15}
                textAnchor="middle"
                className={styles.rankCategory}
                aria-hidden="true"
              >
                {lines.map((line, index) => (
                  <tspan key={line} x={x(category)} dy={index ? 15 : 0}>
                    {line}
                  </tspan>
                ))}
              </text>
            )
          })}
          {profiles.map(({ id }, position) => (
            <line
              key={id}
              x1={labels}
              x2={both ? end : last}
              y1={y(position)}
              y2={y(position)}
              className={styles.rankGuide}
              aria-hidden="true"
            />
          ))}
          {drawn.map(({ profile, color, dash, places }) => {
            const [name, model] = label(profile)
            return (
              <g
                key={profile.id}
                className={cn(
                  styles.rankSeries,
                  dimmed(profile) && styles.rankDimmed,
                )}
                style={{ color }}
                aria-hidden="true"
                onMouseEnter={() => setHovered(profile.id)}
                onClick={() => toggle(profile.id)}
              >
                <path
                  d={path(places)}
                  className={styles.rankLine}
                  strokeDasharray={dash}
                />
                {/* A wider band around the line, easier to point at. */}
                <path d={path(places)} className={styles.rankHit} />
                {places.map(({ position, rank, score }, category) => (
                  <g key={category} className={styles.rankNode}>
                    <circle cx={x(category)} cy={y(position)} r="9.5" />
                    <text
                      x={x(category)}
                      y={y(position)}
                      textAnchor="middle"
                      dominantBaseline="central"
                    >
                      {rank ?? "–"}
                    </text>
                    <title>
                      {`${CATEGORY_DETAILS[category].name}: ${name}, ${model}, rank ${rank ?? "none"} of ${profiles.length}, ${scoreText(score)}`}
                    </title>
                  </g>
                ))}
              </g>
            )
          })}
          {series.map((item) => (
            <g
              key={item.profile.id}
              role="button"
              tabIndex={0}
              aria-pressed={pinned === item.profile.id}
              aria-label={[
                label(item.profile).join(", "),
                shownReasoning(item.profile),
              ]
                .filter(Boolean)
                .join(" ")}
              className={cn(
                styles.rankKey,
                dimmed(item.profile) && styles.rankDimmed,
              )}
              onMouseEnter={() => setHovered(item.profile.id)}
              onFocus={() => setHovered(item.profile.id)}
              onBlur={() => setHovered(null)}
              onClick={() => toggle(item.profile.id)}
              onKeyDown={(event) => {
                if (!["Enter", " "].includes(event.key)) return
                event.preventDefault()
                toggle(item.profile.id)
              }}
            >
              {sideLabel(item, "left")}
            </g>
          ))}
          {both &&
            series.map((item) => (
              <g
                key={item.profile.id}
                aria-hidden="true"
                className={cn(
                  styles.rankKey,
                  dimmed(item.profile) && styles.rankDimmed,
                )}
                onMouseEnter={() => setHovered(item.profile.id)}
                onClick={() => toggle(item.profile.id)}
              >
                {sideLabel(item, "right")}
              </g>
            ))}
        </svg>
      </div>
    </div>
  )
}
