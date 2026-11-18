import { finite } from "./benchmark-model"

// Plot geometry in pixels: a `width` × `height` box whose plot starts 80px
// in, as the overall score's does (the panel's 24px padding and a 56px axis
// gutter holding the vertical score title and the ticks), and runs to the
// panel's padding on the right, where the last duration tick's label sits.
export const plotArea = (width, height) => ({
  width,
  height,
  left: 80,
  right: width - 24,
  top: 24,
  bottom: height - 64,
})

// The configurations a point can place: a score on the scale and a
// duration.
export const scatterProfiles = (profiles) =>
  profiles.filter(
    (profile) =>
      finite(profile.score) &&
      profile.score >= 0 &&
      profile.score <= 100 &&
      finite(profile.duration) &&
      profile.duration >= 0,
  )

// Calculated from every configuration, not the current filter, so the time
// scale stays fixed. Headroom keeps the slowest marker inside the plot.
export function durationAxis(profiles) {
  const maximum = Math.max(
    1000,
    ...scatterProfiles(profiles).map((profile) => profile.duration),
  )
  const roughStep = (maximum * 1.08) / 5
  const magnitude = 10 ** Math.floor(Math.log10(roughStep))
  const step =
    [1, 2, 3, 5, 6, 10].find((n) => n * magnitude >= roughStep) * magnitude
  return {
    max: step * 5,
    ticks: Array.from({ length: 6 }, (_, i) => i * step),
  }
}

export function scoreAxis(profiles) {
  const scores = scatterProfiles(profiles).map((profile) => profile.score)
  const lowest = scores.length ? Math.min(...scores) - 1 : 0
  const step = lowest >= 70 ? 5 : lowest >= 40 ? 10 : 20
  // The axis starts on a step, so its ticks are evenly spaced.
  const min = Math.max(0, Math.floor(lowest / step) * step)
  const ticks = []
  for (let tick = min; tick <= 100; tick += step) ticks.push(tick)
  return { min, ticks }
}

// Place model labels beside the icons after reserving space for every icon,
// so a long label never covers another marker: to the right or left, then
// just above or below, centered or flush with one of the icon's edges. A
// label that finds no room next to its icon is left out (`labelHidden`) and
// shows on hover or focus, rather than floating where it could name a
// neighbour.
export function labelPoints(points, plot) {
  const occupied = points.map((p) => ({
    left: p.bubbleX - p.halfX,
    right: p.bubbleX + p.halfX,
    top: p.bubbleY - p.halfY,
    bottom: p.bubbleY + p.halfY,
  }))
  const intersects = (a, b) =>
    a.left < b.right && a.right > b.left && a.top < b.bottom && a.bottom > b.top
  const spots = (point) => {
    // A 12px Roboto Mono character is 7.2px wide.
    const width = point.label.length * 7.2
    const candidates = []
    for (const dy of [0, -16, 16]) {
      for (const left of [
        point.bubbleX + point.halfX + 4,
        point.bubbleX - point.halfX - 4 - width,
      ]) {
        candidates.push({
          left,
          right: left + width,
          top: point.bubbleY - 7 + dy,
          bottom: point.bubbleY + 7 + dy,
          level: dy === 0,
        })
      }
    }
    for (const top of [
      point.bubbleY - point.halfY - 15,
      point.bubbleY + point.halfY + 1,
    ])
      for (const left of [
        point.bubbleX - width / 2,
        point.bubbleX - point.halfX,
        point.bubbleX + point.halfX - width,
      ])
        candidates.push({ left, right: left + width, top, bottom: top + 14 })
    const inside = candidates.filter(
      (b) =>
        b.left >= plot.left + 2 &&
        b.right <= plot.right - 2 &&
        b.top >= plot.top - 16 &&
        b.bottom <= plot.bottom + 12,
    )
    return { candidates, inside }
  }
  const free = (box) => !occupied.some((r) => intersects(box, r))
  // The most cramped icons choose first: fewest spots clear of every icon.
  const order = points
    .map((point, index) => ({
      index,
      room: spots(point).inside.filter(free).length,
    }))
    .sort((a, b) => a.room - b.room)
  const placed = []
  for (const { index } of order) {
    const point = points[index]
    const { candidates, inside } = spots(point)
    const box = inside.find(free)
    if (box)
      occupied.push({
        left: box.left - 4,
        right: box.right + 4,
        top: box.top - 3,
        bottom: box.bottom + 3,
      })
    const shown = box || inside[0] || candidates[0]
    // A label off its icon's level gets a line from the icon's center to
    // the label's nearest edge; the icon drawn over it hides the inner part.
    const clamp = (value, min, max) => Math.min(Math.max(value, min), max)
    placed[index] = {
      ...point,
      labelX: shown.left,
      labelY: shown.top + 11,
      labelHidden: !box,
      labelLink: shown.level
        ? null
        : {
            x: clamp(point.bubbleX, shown.left, shown.right),
            y: clamp(point.bubbleY, shown.top + 2, shown.bottom - 1),
          },
    }
  }
  return placed
}
