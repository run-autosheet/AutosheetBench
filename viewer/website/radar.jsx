import cn from "classnames"
import { CATEGORY_DETAILS, finite, scoreText } from "./benchmark-model"
import styles from "./benchmark.module.css"

// The eight categories' radar, shared by head to head and the capability
// table's radar view: rings at 25, 50, 75 and 100, a spoke per category
// with its name, and one series per configuration.
export const RADAR_VIEWBOX = "-24 0 528 450"
const CENTER = [240, 215]
const RADIUS = 146
const point = (index, value, radius = RADIUS) => {
  const angle = (index * Math.PI) / 4 - Math.PI / 2
  return [
    CENTER[0] + (Math.cos(angle) * radius * value) / 100,
    CENTER[1] + (Math.sin(angle) * radius * value) / 100,
  ]
}
const polygon = (values) =>
  values.map((value, index) => point(index, value).join(",")).join(" ")
// Which of the series a point in the radar's coordinates is on: the one
// whose outline passes nearest, so the pointer picks a shape anywhere on
// the chart, across fills that overlap, without aiming at a line. Only
// within the 100 ring; null outside it or with no series.
export const nearestSeries = (series, [x, y]) => {
  if (Math.hypot(x - CENTER[0], y - CENTER[1]) > RADIUS + 12) return null
  const toSegment = ([x1, y1], [x2, y2]) => {
    const [dx, dy] = [x2 - x1, y2 - y1]
    const t = Math.max(
      0,
      Math.min(1, ((x - x1) * dx + (y - y1) * dy) / (dx * dx + dy * dy || 1)),
    )
    return Math.hypot(x - (x1 + t * dx), y - (y1 + t * dy))
  }
  let nearest = null
  let best = Infinity
  series.forEach((values, index) => {
    values.forEach((value, i) => {
      const next = (i + 1) % values.length
      if (!finite(value) || !finite(values[next])) return
      const distance = toSegment(point(i, value), point(next, values[next]))
      if (distance < best) [nearest, best] = [index, distance]
    })
  })
  return nearest
}

// A category's full name in lines of at most 15 characters.
const LINE = 15
export const nameLines = (name) =>
  name.split(" ").reduce((lines, word) => {
    const last = lines.at(-1)
    if (last && `${last} ${word}`.length <= LINE)
      lines[lines.length - 1] = `${last} ${word}`
    else lines.push(word)
    return lines
  }, [])

// The rings, spokes and category names; `counts`, when given, adds each
// category's number of tasks under its name. With `onSelect`, a name that
// `selectable` allows is a button, by pointer or keyboard. `ticks` names the
// four rings, from the center out, where they aren't scores.
const RINGS = [25, 50, 75, 100]
export function RadarAxes({ counts, ticks = RINGS, onSelect, selectable }) {
  return (
    <>
      {RINGS.map((value, ring) => (
        <g key={value}>
          <polygon
            points={polygon(Array(8).fill(value))}
            className={styles.radarGrid}
          />
          <text
            x={CENTER[0] + 6}
            y={CENTER[1] - (RADIUS * value) / 100 - 4}
            className={styles.radarTick}
          >
            {ticks[ring]}
          </text>
        </g>
      ))}
      {CATEGORY_DETAILS.map(({ name }, index) => {
        const end = point(index, 100)
        const text = point(index, 100, 181)
        const lines = nameLines(name)
        // The top label rises, so its count clears the 100 tick; a side
        // label of two lines centers on its spoke.
        const first =
          index === 0
            ? -8 - (lines.length - 1) * 15
            : index === 4
              ? 4
              : 4 - ((lines.length - 1) * 15) / 2
        return (
          <g key={name}>
            <line
              x1={CENTER[0]}
              y1={CENTER[1]}
              x2={end[0]}
              y2={end[1]}
              className={styles.radarGrid}
            />
            <text
              x={text[0]}
              y={text[1]}
              textAnchor={
                index === 0 || index === 4
                  ? "middle"
                  : index < 4
                    ? "start"
                    : "end"
              }
              className={styles.radarLabel}
              {...(onSelect &&
                selectable?.(index) && {
                  role: "button",
                  tabIndex: 0,
                  className: cn(styles.radarLabel, styles.radarLabelLink),
                  "aria-label": `Compare the ${name} tasks`,
                  onClick: () => onSelect(index),
                  onKeyDown: (event) => {
                    if (event.key !== "Enter" && event.key !== " ") return
                    event.preventDefault()
                    onSelect(index)
                  },
                })}
            >
              {lines.map((line, row) => (
                <tspan key={line} x={text[0]} dy={row ? 15 : first}>
                  {line}
                </tspan>
              ))}
              {/* The count on a line of its own, so side labels stay in. */}
              {counts && (
                <tspan x={text[0]} dy="15" className={styles.radarCount}>
                  {`(n=${counts[index]})`}
                </tspan>
              )}
            </text>
          </g>
        )
      })}
    </>
  )
}

// One configuration's eight scores, in `color`, each vertex titled with
// `name`, its category and its score, or with what `title` words for its
// category when the values aren't scores. `dashed` draws head to head's
// second side; `dash` sets any other pattern.
export function RadarSeries({
  values,
  name,
  title,
  color,
  dashed,
  dash,
  className,
}) {
  return (
    <g
      className={cn(styles.radarSeries, dashed && styles.dashed, className)}
      style={{ color }}
    >
      {values.every(finite) ? (
        <polygon points={polygon(values)} strokeDasharray={dash} />
      ) : (
        // With a category missing, draw only the edges between scored
        // neighbors rather than closing the shape through zero.
        values.map((value, index) => {
          const next = (index + 1) % 8
          if (!finite(value) || !finite(values[next])) return null
          const [x1, y1] = point(index, value)
          const [x2, y2] = point(next, values[next])
          return (
            <line
              key={index}
              x1={x1}
              y1={y1}
              x2={x2}
              y2={y2}
              strokeDasharray={dash}
            />
          )
        })
      )}
      {values.map(
        (value, index) =>
          finite(value) && (
            <circle
              key={index}
              cx={point(index, value)[0]}
              cy={point(index, value)[1]}
              r="3.5"
            >
              <title>
                {title
                  ? title(index)
                  : `${name}: ${CATEGORY_DETAILS[index].name} ${scoreText(value)}`}
              </title>
            </circle>
          ),
      )}
    </g>
  )
}
