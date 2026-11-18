"use client"

import { useRef } from "react"
import cn from "classnames"
import { DataTable } from "./chart-data"
import { chartOrder, overallScoreData } from "./chart-rows"
import ChartTools, { ChartBrand } from "./chart-tools"
import PlatformPills, { ConfigurationPills } from "./platform-pills"
import { ConnectorLogo, PlatformIcon, ProductLogo } from "./product-logo"
import {
  KINDS,
  PLATFORMS,
  PLATFORM_LABELS,
  PLATFORM_SHORT,
  PROFILES,
  finite,
  pickConfigurations,
  sortProfiles,
} from "./benchmark-model"
import styles from "./benchmark.module.css"

// Every configuration on both platforms, whatever the toggles show, so the
// page's HTML holds all the results.


const byScore = (profiles) =>
  sortProfiles(profiles, { key: "score", direction: "desc" })

// Each configuration's overall score as a lollipop: its logos on a stem
// from a 50% floor, lower where a score needs it. A sidebar add-on shows
// one logo, an assistant using an MCP shows the assistant's logo and the
// MCP's; branded products take the brand gradient. Columns sit under a
// platform band, then a product-type band, and the last line under each
// names the model the assistant ran. The page holds `configurations`, so
// one toggle picks the configurations of both score charts. A phone turns
// the columns into rows, the stems running across; positions come as
// `--at`, read as a height or a width.
export default function ScoreColumns({
  profiles,
  platform,
  onPlatformChange,
  configurations,
  onConfigurationsChange,
  id,
}) {
  const DATA = overallScoreData(chartOrder(profiles))
  const panel = useRef(null)
  const scored = pickConfigurations(
    profiles.filter((profile) => finite(profile.score)),
    configurations,
  )
  const groups = PLATFORMS.map((platform) => ({
    platform,
    kinds: KINDS.map((kind) => ({
      ...kind,
      columns: byScore(
        scored.filter(
          (profile) =>
            profile.platform === platform && !!profile.connector === kind.mcp,
        ),
      ),
    })).filter((kind) => kind.columns.length),
  })).filter((group) => group.kinds.length)
  const floor = Math.min(
    50,
    Math.floor(Math.min(...scored.map((profile) => profile.score)) / 10) * 10,
  )
  const ticks = Array.from(
    { length: (100 - floor) / 10 + 1 },
    (_, index) => floor + index * 10,
  )
  const position = (score) => `${((score - floor) / (100 - floor)) * 100}%`
  const logos = (profile) => (
    <span
      className={cn(
        styles.scoreColumnLogos,
        profile.connector && styles.pairLogos,
        profile.connector && profile.branded && styles.brandLogos,
      )}
    >
      {/* In a pair, the pill is the assistant logo's light surface. */}
      <ProductLogo
        profile={{ productKey: profile.productKey }}
        size={40}
        bare={!!profile.connector}
      />
      {profile.connector && (
        <>
          <span className={styles.pairPlus} aria-hidden="true">
            +
          </span>
          <ConnectorLogo connector={profile.connector} size={40} />
        </>
      )}
    </span>
  )
  const count = (kinds) =>
    kinds.reduce((total, kind) => total + kind.columns.length, 0)
  const columns = groups.reduce((total, group) => total + count(group.kinds), 0)
  // The toggle already names a single platform, so its header goes; so does
  // the product-type header when that platform has one type (Excel).
  const showPlatform = platform === "all"
  const showKind =
    showPlatform ||
    new Set(scored.map((profile) => !!profile.connector)).size > 1
  return (
    <div id={id} ref={panel} className={styles.leaderboardPanel}>
      <div className={styles.scoreColumnsIntro}>
        <div>
          <h3>Task pass rate</h3>
          {/* What a pass means, for every score on the page. */}
          <a className={styles.scoringLink} href="#snapshot-notes" data-image="hide">
            Saved judgments, without rescoring
          </a>
        </div>
        <div className={styles.chartAside}>
          <div className={styles.chartControls} data-image="hide">
            <PlatformPills platform={platform} onChange={onPlatformChange} />
            <ConfigurationPills
              configurations={configurations}
              onChange={onConfigurationsChange}
            />
            <ChartTools
              anchor={id}
              file="overall-score"
              image={panel}
              data={() => overallScoreData(chartOrder(scored))}
            />
          </div>
          <ChartBrand />
        </div>
      </div>
      <div
        className={styles.scoreColumnsScroll}
        data-image="scroll"
        role="region"
        aria-label="Overall score chart"
        tabIndex={0}
      >
        <div
          className={cn(
            styles.scoreColumns,
            groups.length === 1 && styles.isSinglePlatform,
          )}
          style={{
            "--columns": columns,
            "--bands": `${(showPlatform ? 36 : 0) + (showKind ? 36 : 0)}px`,
          }}
        >
          <div className={styles.scoreColumnsGrid} aria-hidden="true">
            {ticks.map((tick) => (
              <span key={tick} style={{ bottom: position(tick) }}>
                <span>{`${tick}%`}</span>
              </span>
            ))}
          </div>
          <div className={styles.scoreColumnBands}>
            {groups.map(({ platform, kinds }) => (
              <div key={platform} style={{ "--span": count(kinds) }}>
                {showPlatform && (
                  <div className={styles.scoreColumnPlatform}>
                    <PlatformIcon platform={platform} size={40} />
                    {PLATFORM_LABELS[platform]}
                  </div>
                )}
                <div className={styles.scoreColumnKinds}>
                  {kinds.map((kind) => (
                    <div
                      key={kind.label}
                      style={{ "--span": kind.columns.length }}
                    >
                      {showKind && (
                        <div className={styles.scoreColumnKind}>
                          {kind.label}
                        </div>
                      )}
                      <ol
                        className={styles.scoreColumnsList}
                        aria-label={`${PLATFORM_LABELS[platform]}, ${kind.label}`}
                      >
                        {kind.columns.map((profile) => (
                          <li key={profile.id}>
                            <div className={styles.scoreColumnPlot}>
                              <span
                                className={styles.plotGrid}
                                aria-hidden="true"
                              >
                                {ticks.map((tick) => (
                                  <span
                                    key={tick}
                                    style={{ "--at": position(tick) }}
                                  />
                                ))}
                              </span>
                              <span
                                className={cn(
                                  styles.scoreStem,
                                  profile.branded && styles.brandStem,
                                )}
                                style={{ "--at": position(profile.score) }}
                              />
                              <span
                                className={styles.scoreDot}
                                style={{ "--at": position(profile.score) }}
                              >
                                <span className={styles.scoreColumnValue}>
                                  {`${profile.score.toFixed(1)}%`}
                                </span>
                                {logos(profile)}
                              </span>
                            </div>
                            <div className={styles.scoreColumnLabel}>
                              <strong>{profile.assistant}</strong>
                              {profile.connector ? (
                                <span>
                                  {`+ ${profile.mcp}`}
                                  <br />
                                  {/* Spaced, for a phone's row, which
                                      drops the break. */}
                                  {" MCP"}
                                </span>
                              ) : (
                                <span>{`in ${PLATFORM_SHORT[profile.platform]}`}</span>
                              )}
                              {/* A phone's rows give the score here, the
                                  stem being too short to carry it. */}
                              <b aria-hidden="true">
                                {`${profile.score.toFixed(1)}%`}
                              </b>
                              <em>{profile.modelName}</em>
                            </div>
                          </li>
                        ))}
                      </ol>
                    </div>
                  ))}
                </div>
              </div>
            ))}
          </div>
        </div>
      </div>
      <DataTable
        caption="Overall score: every configuration, on both platforms"
        data={DATA}
      />
    </div>
  )
}
