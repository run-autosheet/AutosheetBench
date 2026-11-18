"use client"

import { useEffect, useId, useRef, useState } from "react"
import cn from "classnames"
import { domToBlob } from "modern-screenshot"
import { BENCHMARK_NAME } from "./benchmark-model"
import { toCsv } from "./chart-rows"
import { Check, ImageDown, Link, Share, Spreadsheet } from "./icons"
import styles from "./benchmark.module.css"

// Saves `blob` as `name` through a link the browser follows at once.
const download = (name, blob) => {
  const url = URL.createObjectURL(blob)
  const link = document.createElement("a")
  link.href = url
  link.download = name
  document.body.append(link)
  link.click()
  link.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}

// Copies `text`, through a hidden field where the clipboard is refused,
// and says whether it did; where both fail, it offers the text to copy by
// hand.
async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text)
    return true
  } catch {
    const field = document.createElement("textarea")
    field.value = text
    field.setAttribute("readonly", "")
    field.className = styles.visuallyHidden
    document.body.append(field)
    field.select()
    const copied = document.execCommand("copy")
    field.remove()
    if (!copied) window.prompt("Copy this link", text)
    return copied
  }
}

// The chart's card as a PNG at twice its size, without what
// `data-image="hide"` marks: the toggles, these tools, links and the data
// table. It is drawn from a copy laid out out of sight, so the chart on the
// page doesn't move, as wide as the parts marked `data-image="scroll"`
// need to show all of themselves: a chart that scrolls sideways on a phone
// comes out whole. The library loads on the first click, not with the page.
async function chartImage(node) {
  const box = node.getBoundingClientRect()
  const width = Math.max(
    box.width,
    ...[...node.querySelectorAll('[data-image="scroll"]')].map((part) => {
      const { left, right } = part.getBoundingClientRect()
      return left - box.left + part.scrollWidth + (box.right - right)
    }),
  )
  // Without ids, which the page's own chart holds, and shown from its
  // start, so without the table's scrolled edge.
  const copy = node.cloneNode(true)
  for (const element of [copy, ...copy.querySelectorAll("[id]")])
    element.removeAttribute("id")
  for (const element of copy.querySelectorAll("[data-scrolled]"))
    element.removeAttribute("data-scrolled")
  copy.classList.add(styles.chartImage)
  copy.style.width = `${Math.ceil(width)}px`
  copy.setAttribute("aria-hidden", "true")
  copy.inert = true
  node.after(copy)
  try {
    return await domToBlob(copy, {
      scale: 2,
      backgroundColor: getComputedStyle(node).backgroundColor,
      // In the image, the copy sits at its origin.
      style: { position: "static" },
      filter: (element) => element.dataset?.image !== "hide",
    })
  } finally {
    copy.remove()
  }
}

// The benchmark's name on each chart, as Artificial Analysis signs its
// own: the brand's icon and name, quiet at the chart's top right, and kept
// in the chart's image. Screen readers skip it; the page's title says it.
// The icon has a quiet edge of its own, so it holds on a dark card
// without the products' green ring.
export function ChartBrand({ className }) {
  return (
    <div className={cn(styles.chartBrand, className)} aria-hidden="true">
      <svg width="22" height="22" viewBox="0 0 16 16" focusable="false">
        <image
          href="/icons/favicon.svg"
          width="16"
          height="16"
          className={styles.brandIcon}
        />
        <rect
          className={styles.brandIconEdge}
          width="16"
          height="16"
          rx="2.88"
        />
      </svg>
      {BENCHMARK_NAME}
    </div>
  )
}

// A chart's own actions behind one share button, so they take a single
// icon's room beside the toggles: a menu of the link, the image and the
// data. `anchor` is the id a link to the chart jumps to; `image` holds the
// chart's card; `data` gives the rows the chart shows, read only when asked
// for; `file` names what is saved.
export default function ChartTools({ anchor, file, image, data }) {
  const [open, setOpen] = useState(false)
  const menu = useId()
  const [busy, setBusy] = useState(false)
  const [copied, setCopied] = useState(false)
  const root = useRef(null)
  const trigger = useRef(null)
  const name = `${BENCHMARK_NAME.toLowerCase()}-${file}`
  // A click outside or Escape closes the menu; Escape gives the focus back
  // to the share button.
  useEffect(() => {
    if (!open) return
    const outside = (event) => {
      if (!root.current?.contains(event.target)) setOpen(false)
    }
    const escape = (event) => {
      if (event.key !== "Escape") return
      setOpen(false)
      trigger.current?.focus()
    }
    document.addEventListener("pointerdown", outside)
    document.addEventListener("keydown", escape)
    return () => {
      document.removeEventListener("pointerdown", outside)
      document.removeEventListener("keydown", escape)
    }
  }, [open])
  // The menu stays open on a copied link, which it confirms.
  const copyLink = async () => {
    const { origin, pathname } = window.location
    if (!(await copyText(`${origin}${pathname}#${anchor}`))) return
    setCopied(true)
    setTimeout(() => setCopied(false), 2000)
  }
  // Busy rather than disabled, which would drop a keyboard user's focus.
  const saveImage = async () => {
    if (busy) return
    setBusy(true)
    try {
      download(`${name}.png`, await chartImage(image.current))
      setOpen(false)
    } catch (error) {
      console.error("The chart image could not be made.", error)
    } finally {
      setBusy(false)
    }
  }
  const saveData = () => {
    download(
      `${name}.csv`,
      new Blob([toCsv(data())], { type: "text/csv;charset=utf-8" }),
    )
    setOpen(false)
  }
  return (
    <div ref={root} className={styles.chartTools} data-image="hide">
      <div
        className={cn(styles.pills, styles.platformPills, styles.chartPills)}
      >
        <button
          ref={trigger}
          type="button"
          aria-label="Share and download"
          title="Share and download"
          // A group of buttons, not a menu: no aria-haspopup.
          aria-expanded={open}
          aria-controls={open ? menu : undefined}
          onClick={() => setOpen(!open)}
        >
          <Share size={15} />
        </button>
      </div>
      {open && (
        <div
          id={menu}
          className={styles.shareMenu}
          role="group"
          aria-label="Share and download"
        >
          <button type="button" onClick={copyLink}>
            {copied ? <Check size={15} /> : <Link size={15} />}
            {copied ? "Link copied" : "Copy link to this chart"}
          </button>
          <button
            type="button"
            aria-busy={busy}
            aria-disabled={busy}
            onClick={saveImage}
          >
            <ImageDown size={15} />
            {busy ? "Preparing the image…" : "Download as image"}
          </button>
          <button type="button" onClick={saveData}>
            <Spreadsheet size={15} />
            Download data as CSV
          </button>
        </div>
      )}
      <span className={styles.visuallyHidden} role="status">
        {copied ? "Link copied" : ""}
      </span>
    </div>
  )
}
