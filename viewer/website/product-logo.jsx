import cn from "classnames"
import { PLATFORM_LABELS } from "./platform-labels"
import { shownReasoning } from "./model-labels"
import styles from "./benchmark.module.css"

const LOGOS = {
  "gpt-for-excel": "/icons/favicon.svg",
  "gpt-for-sheets": "/icons/favicon.svg",
  chatgpt: "/icons/icon-openai.svg",
  claude: "/icons/icon-claude.svg",
  copilot: "/icons/icon-copilot.svg",
  gemini: "/icons/icon-gemini.svg",
}
const CONNECTOR_LOGOS = {
  autosheet: "/icons/favicon.svg",
  "google-sheets": "/icons/icon-google-sheets.svg",
  "google-drive": "/icons/icon-google-sheets.svg",
}
// Third-party logos sit on a disk so the dark OpenAI mark stays readable on
// dark rows; the brand icon is a filled square and needs none.
const DISKED = new Set(["chatgpt", "claude", "copilot", "gemini"])
const BRAND_ICON = "/icons/favicon.svg"
// The brand square takes the header logo's rounded corners.
const iconClass = (href) => (href === BRAND_ICON ? styles.brandIcon : undefined)

// Product and connector colors, read from tokens on the page root so both
// themes keep the lines and markers readable.
const COLOR_KEYS = {
  "gpt-for-excel": "brand",
  "gpt-for-sheets": "brand",
  claude: "claude",
  chatgpt: "chatgpt",
  copilot: "copilot",
  gemini: "gemini",
}
const colorKey = (profile) => COLOR_KEYS[profile?.productKey] || "other"

// Many configurations on one chart: a product's configurations alternate
// its color and its alternate, solid, then dashed, dotted and dash-dotted,
// so eight of one product still read apart.
const DASHES = [undefined, "5 3", "1.5 3", "8 3 1.5 3"]
export const seriesStyles = (profiles) => {
  const seen = {}
  return profiles.map((profile) => {
    const key = colorKey(profile)
    const n = (seen[key] = (seen[key] ?? -1) + 1)
    return {
      color: `var(--benchmark-${key}${n % 2 ? "-alt" : ""})`,
      dash: DASHES[Math.floor(n / 2) % DASHES.length],
    }
  })
}

// An SVG group centered on (x, y). A connector configuration shows the host's
// logo with the connector's logo as a badge. `bare` drops the disk, for a
// logo that already sits on a light surface, and keeps the logo's size.
export function ProductLogoMark({
  productKey,
  connector,
  x,
  y,
  size = 32,
  bare = false,
}) {
  const badge = CONNECTOR_LOGOS[connector]
  if (badge) {
    const badgeSize = size * 0.46
    const pad = size * 0.035
    return (
      <g transform={`translate(${x - size / 2} ${y - size / 2})`}>
        <ProductLogoMark
          productKey={productKey}
          x={size * 0.38}
          y={size * 0.38}
          size={size * 0.72}
        />
        <g transform={`translate(${size * 0.53} ${size * 0.53})`}>
          <rect
            className={styles.logoBadge}
            x={-pad}
            y={-pad}
            width={badgeSize + pad * 2}
            height={badgeSize + pad * 2}
            rx={size * 0.08}
            strokeWidth={size * 0.02}
          />
          <image
            href={badge}
            className={iconClass(badge)}
            width={badgeSize}
            height={badgeSize}
          />
        </g>
      </g>
    )
  }
  const logo = LOGOS[productKey]
  return (
    <g transform={`translate(${x - size / 2} ${y - size / 2})`}>
      {DISKED.has(productKey) && !bare && (
        <circle
          className={styles.logoDisk}
          cx={size / 2}
          cy={size / 2}
          r={size * 0.64}
        />
      )}
      {logo ? (
        <image
          href={logo}
          className={iconClass(logo)}
          x={DISKED.has(productKey) ? size * 0.1 : 0}
          y={DISKED.has(productKey) ? size * 0.1 : 0}
          width={DISKED.has(productKey) ? size * 0.8 : size}
          height={DISKED.has(productKey) ? size * 0.8 : size}
        />
      ) : (
        <text
          className={styles.logoInitial}
          x={size / 2}
          y={size / 2}
          fontSize={size * 0.36}
          textAnchor="middle"
          dominantBaseline="central"
        >
          {String(productKey || "?")
            .slice(0, 1)
            .toUpperCase()}
        </text>
      )}
      {logo === BRAND_ICON && (
        <rect
          className={styles.brandIconEdge}
          width={size}
          height={size}
          rx={size * 0.18}
        />
      )}
    </g>
  )
}

// An assistant and its MCP side by side on a pill, with a plus between
// them, as in the overall score chart. Sizes are fractions of `size`, the
// nominal mark size: the bare assistant logo is drawn at 0.8 as on its
// disk, the MCP icon at 1.
export const PAIR_MARK = { width: 3.1, height: 1.5 }
export function ProductPairMark({
  productKey,
  connector,
  x,
  y,
  size,
  branded,
}) {
  const pad = size * 0.25
  const gap = size * 0.2
  const plus = size * 0.4
  const width = size * PAIR_MARK.width
  const height = size * PAIR_MARK.height
  const left = x - width / 2
  const icon = CONNECTOR_LOGOS[connector]
  return (
    <g>
      <rect
        className={cn(styles.pairPill, branded && styles.brandPill)}
        x={left}
        y={y - height / 2}
        width={width}
        height={height}
        rx={height / 2}
      />
      <ProductLogoMark
        productKey={productKey}
        x={left + pad + size * 0.4}
        y={y}
        size={size}
        bare
      />
      <text
        className={styles.pairPlusMark}
        x={left + pad + size * 0.8 + gap + plus / 2}
        y={y}
        fontSize={size * 0.65}
        textAnchor="middle"
        dominantBaseline="central"
      >
        +
      </text>
      <image
        href={icon}
        className={iconClass(icon)}
        x={left + pad + size * 0.8 + gap * 2 + plus}
        y={y - size / 2}
        width={size}
        height={size}
      />
    </g>
  )
}

export const ProductLogo = ({ profile, size = 32, className, bare }) => (
  <svg
    className={className}
    viewBox="0 0 40 40"
    width={size}
    height={size}
    aria-hidden="true"
    focusable="false"
  >
    <ProductLogoMark
      productKey={profile?.productKey}
      connector={profile?.connector}
      x={20}
      y={20}
      size={25}
      bare={bare}
    />
  </svg>
)

const PLATFORM_ICONS = {
  microsoft_excel: "/icons/icon-microsoft-excel.svg",
  google_sheets: "/icons/icon-google-sheets.svg",
}

// A square icon drawn like the product logos: 25 units in a 40-unit box.
const IconImage = ({ href, size, className }) => (
  <svg
    className={className}
    viewBox="0 0 40 40"
    width={size}
    height={size}
    aria-hidden="true"
    focusable="false"
  >
    <image
      href={href}
      className={iconClass(href)}
      x="7.5"
      y="7.5"
      width="25"
      height="25"
    />
  </svg>
)

export const ConnectorLogo = ({ connector, size = 32 }) => (
  <IconImage href={CONNECTOR_LOGOS[connector]} size={size} />
)

// A platform icon on a small white tile, centered on (x, y), to badge a
// scatter marker with its spreadsheet.
export const PlatformBadgeMark = ({ platform, x, y, size }) => (
  <g transform={`translate(${x - size / 2} ${y - size / 2})`}>
    <rect
      className={styles.logoBadge}
      x={-size * 0.12}
      y={-size * 0.12}
      width={size * 1.24}
      height={size * 1.24}
      rx={size * 0.24}
    />
    <image href={PLATFORM_ICONS[platform]} width={size} height={size} />
  </g>
)

export const PlatformIcon = ({ platform, size = 24, className }) => (
  <IconImage
    href={PLATFORM_ICONS[platform]}
    size={size}
    className={className}
  />
)

// The platform, the assistant, its connection to the spreadsheet and its
// model, each under its own heading, so "Claude" never has to stand for both.
export const ConfigurationSlots = ({ profile }) => (
  <dl className={styles.configSlots}>
    <div>
      <dt>Platform</dt>
      <dd>{PLATFORM_LABELS[profile.platform]}</dd>
    </div>
    <div>
      <dt>Assistant</dt>
      <dd>{profile.assistant}</dd>
    </div>
    <div>
      <dt>Connection</dt>
      <dd>{profile.connection}</dd>
    </div>
    <div>
      <dt>Model</dt>
      <dd>
        {[profile.modelName, shownReasoning(profile)]
          .filter(Boolean)
          .join(" · ")}
        {profile.mcpModel && <small>{`MCP runs ${profile.mcpModel}`}</small>}
      </dd>
    </div>
  </dl>
)

