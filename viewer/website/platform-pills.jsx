import { createContext, useContext } from "react"
import {
  CONFIGURATIONS,
  PLATFORMS,
  PLATFORM_SHORT,
  PRODUCT_TYPES,
  PROFILES,
  inView,
} from "./benchmark-model"
import { Pills } from "./pills"
import styles from "./benchmark.module.css"

// Excel first, as the default view, whatever order the results arrive in.
// Short words, so each toggle reads at a glance.
const PLATFORM_CHOICES = [
  ...Object.keys(PLATFORM_SHORT)
    .filter((value) => PLATFORMS.includes(value))
    .map((value) => ({ value, label: PLATFORM_SHORT[value] })),
  { value: "all", label: "All" },
]

// The page holds the product type, as `[value, setValue]`; every copy of
// the platform toggle is followed by it.
export const ProductTypeContext = createContext(["all", () => {}])
export const useProductType = () => useContext(ProductTypeContext)

// `types={false}` shows the platform alone, for a view that lists every
// type.
export default function PlatformPills({ platform, onChange, types = true }) {
  const [productType, setProductType] = useProductType()
  // A type the platform has no product of moves to the first platform that
  // has one: MCP on Excel switches to Google Sheets.
  const chooseType = (value) => {
    if (!PROFILES.some((profile) => inView(profile, platform, value)))
      onChange(
        PLATFORM_CHOICES.find((choice) =>
          PROFILES.some((profile) => inView(profile, choice.value, value)),
        ).value,
      )
    setProductType(value)
  }
  return (
    <div className={styles.platformChoice}>
      <Pills
        choices={PLATFORM_CHOICES}
        value={platform}
        onChange={onChange}
        label="Spreadsheet platform"
      />
      {types && (
        <Pills
          choices={PRODUCT_TYPES}
          value={productType}
          onChange={chooseType}
          label="Product type"
        />
      )}
    </div>
  )
}

// Which configurations the score charts plot: each product's best model,
// or every one tested.
export function ConfigurationPills({ configurations, onChange }) {
  return (
    <Pills
      choices={CONFIGURATIONS}
      value={configurations}
      onChange={onChange}
      label="Models shown"
      className={styles.modelsPills}
    />
  )
}
