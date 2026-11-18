import { useLayoutEffect, useRef } from "react"
import cn from "classnames"
import styles from "./benchmark.module.css"

// A row of pills the page holds the value of, so every copy of the toggle
// (one per chart, the table and the gallery) moves together.
export function Pills({ choices, value, onChange, label, className }) {
  const group = useRef(null)
  const clickedTop = useRef(null)
  // Switching resizes the charts above this copy; scroll by the difference
  // so the toggle that was clicked stays under the pointer.
  useLayoutEffect(() => {
    if (clickedTop.current === null || !group.current) return
    const moved = group.current.getBoundingClientRect().top - clickedTop.current
    clickedTop.current = null
    // Instant: the site scrolls smoothly, which would slide it back.
    if (moved) window.scrollBy({ top: moved, behavior: "instant" })
  }, [value])
  return (
    <div
      ref={group}
      className={cn(
        styles.pills,
        styles.platformPills,
        styles.chartPills,
        className,
      )}
      role="group"
      aria-label={label}
    >
      {choices.map((choice) => (
        <button
          type="button"
          key={choice.value}
          aria-pressed={value === choice.value}
          onClick={() => {
            if (choice.value === value) return
            clickedTop.current = group.current.getBoundingClientRect().top
            onChange(choice.value)
          }}
        >
          {choice.label}
        </button>
      ))}
    </div>
  )
}
