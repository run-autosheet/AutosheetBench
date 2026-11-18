import { memo } from "react"
import styles from "./benchmark.module.css"

// A chart's data as a table, out of sight: screen readers and machines that
// read the page's HTML get every value the chart draws. The wrapper clips
// the table, which a size of its own wouldn't. Memoized, as most tables
// hold the module's fixed rows the toggles would otherwise rebuild.
export const DataTable = memo(function DataTable({ caption, data }) {
  return (
    <div className={styles.visuallyHidden} data-image="hide">
      <table>
        <caption>{caption}</caption>
        <thead>
          <tr>
            {data.columns.map((column) => (
              <th key={column} scope="col">
                {column}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {data.rows.map((row, index) => (
            <tr key={index}>
              {row.map((value, column) =>
                column ? (
                  <td key={column}>{value ?? "No data"}</td>
                ) : (
                  <th key={column} scope="row">
                    {value}
                  </th>
                ),
              )}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )
})
