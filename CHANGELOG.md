# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project follows [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [1.4.1] - 2026-09-17

### Fixed

- Shrink a wrapped `normAutofit` body by the height it overflows rather than by the width of an
  unwrapped line. Such a box was measured twice — once as it wraps, once with white-space forced
  to `nowrap` — and the second measurement won whenever the wrapped text fitted the width exactly
  and overflowed only the height. The box was then scaled by its width against that of a single
  unwrapped line, a width wrapped text never occupies. In a Chinese lecture deck three paragraphs
  at 18pt overflowed a 1104x486 box by 12%, the unwrapped line ran to 2514px, and the body
  rendered at roughly 7.9pt beneath a 24pt heading instead of at the 0.89 the height overflow
  called for. PowerPoint keeps paragraphs wrapping at the box width while `normAutofit` shrinks
  them, so the unwrapped width is never the right reference; `wrap="none"` bodies, `spAutoFit`,
  and text that genuinely overflows horizontally once wrapped are unaffected.

## [1.4.0] - 2026-09-15

### Added

- Render WMF pictures instead of a grey "Unsupported format: WMF" box that wiped out a whole
  region of the slide. WMF is the 16-bit ancestor of EMF and draws with the same GDI model, so
  the shared state — pens, brushes, stock objects, map modes, COLORREF, SVG serialization — now
  lives in `src/utils/gdi.ts` and each format keeps only its own record decoding. Verified
  against 190 real clip-art files shipped with WPS Office.
- Draw the text records of a WMF (`TextOut`/`ExtTextOut`) with fonts, charsets, colour,
  alignment and per-glyph `Dx` spacing. A MathType equation pasted into a deck arrives as an OLE
  frame whose only preview is a WMF drawn entirely with text — the curly braces themselves being
  Symbol font slots — so such a picture recovered no geometry at all and left a hole in the
  slide. Text bytes carry no encoding of their own and are read through the selected font's
  `CharSet`, DBCS code pages included; Symbol slots map to Unicode through the Adobe Symbol
  table. EMF text output stays unsupported.
- Render `ofPieChart` (pie of pie and bar of pie), which previously fell through to the
  unsupported branch and replaced the whole plot with a line of small text. The `c:splitType`
  rule that decides which points move to the secondary plot — with its `pos`/`val`/`percent`/
  `cust` variants — lives in `chart/ofPie.ts` as pure functions, because getting it wrong still
  draws a plausible chart that describes different data.
- Render chartex (`cx:chartSpace`) waterfall and funnel charts. Office 2016's second chart
  format shares nothing with `c:chartSpace` but the word "chart", so every one of its layouts
  replaced the plot with a notice. The remaining chartex layouts (treemap, sunburst, box &
  whisker, histogram, Pareto, region map) still report an unsupported-layout notice rather than
  being drawn as a chart type whose numbers mean something else.
- Render the `textCircle` WordArt warp alongside the two arch warps, transcribed from ECMA-376's
  `presetTextWarpDefinitions.xml`. Only three of the 40 presets declare a single `<path>` and so
  can be expressed as an SVG `<textPath>`; the other 37 are deformations between two envelopes
  and continue to render as ordinary unwarped text.

### Fixed

- Keep the WMF object table's free slots distinct from slots held by an object the converter
  cannot draw with. Both were `null`, so a later brush reused a font's slot and shifted every
  subsequent `SELECTOBJECT` index.
- Frame a WMF from the window in force when drawing starts rather than the first window
  declared, and never mistake a device context's default 1x1 extent for a real one. Real clip
  art opens with a placeholder window, and the scalable map modes divided such drawings away:
  correct framing went from 41/190 to 189/190 of the clip-art corpus.
- Give an `ofPieChart`'s aggregate slice the first colour its pie has not used. It indexed past
  the end of the palette, which wraps, so a six-point series on a six-entry theme palette left
  two slices and two legend entries indistinguishable. A bar of pie's column is also pinned to
  its own stack total instead of being scaled against the pie's values, which had collapsed it
  to a fifth of the plot height.

### Changed

- Corrected stale "not supported" claims in the documentation. Pattern fills, shadow/reflection/
  glow effects, combo charts, secondary axes and OLE static previews had all been implemented for
  some time, so the list was steering work toward re-implementing what already existed.

## [1.3.1] - 2026-09-11

### Fixed

- Arc warped text (`textArchUp`/`textArchDown`) along the baseline its `prstTxWarp` geometry
  actually describes. The baseline was approximated with a fixed shallow quadratic that ignored
  the preset's `adj` angle and sat at 36%/66% of the box height instead of on the inscribed
  ellipse, so once each shape's rotation was applied the labels drifted off the artwork they
  annotate. Derived from ECMA-376 `presetTextWarpDefinitions.xml`, emitting real elliptical arcs,
  and the run is now placed by the paragraph's `algn` rather than always centred.
- Load the pictures a chart paints its series with when `lazyMedia` is enabled. A series
  `a:blipFill` is resolved while the ECharts option is assembled, which is synchronous, so under
  `lazyMedia` the bytes were still behind the media resolver, the lookup missed, and the series was
  painted with a palette colour instead of its artwork — silently, with no error. Chart image
  parts are now resolved before the first render; slide media stays lazy.
- Cascade the colour map slide -> layout -> master. A slide's `<a:masterClrMapping/>` means
  "inherit", not "jump to the master `colorMap`", and PowerPoint writes it on nearly every slide,
  so every layout-level `overrideClrMapping` was unreachable: a deck whose layout flips the map
  back to light rendered with the master's dark mapping, white-on-dark where PowerPoint shows
  dark-on-white.
- Recognize the abbreviated ECMA-376 `ST_PresetColorVal` spellings. The preset table held only the
  CSS names (`darkBlue`, `lightGreen`, `mediumPurple`), so all 43 abbreviated forms (`dkBlue`,
  `ltGreen`, `medPurple`, ...) missed it and fell through to black. An unrecognized `prstClr` name
  now falls back to neutral gray, which is distinguishable from a deliberate `val="black"`.

## [1.3.0] - 2026-09-10

### Added

- Added opt-in browser rendering for licensed EOT/MTX fonts embedded in PowerPoint files, with
  bounded face, byte, and processing budgets plus host-font fallback for rejected faces. Enable it
  with the `embeddedFonts` option on `PptxViewer` or headless `renderSlide()`.
- Added `fontFaces` options to `PptxViewer` and headless `renderSlide()` so host applications can
  register missing regular/bold font data before PowerPoint text layout is measured.
- Added a 12-case CJK native-oracle matrix for wrap, autofit, line/paragraph spacing, adjacent
  run spacing, and parent-shape layout, with tracked coverage/font metadata and ignored binaries.
- Added optional local font profiles and per-evaluation provenance for PPTX/ground-truth/font
  hashes, renderer Git state, and the actual browser version.
- Added vector EMF rendering: plain GDI path drawings inside EMF pictures are converted to SVG,
  so decorative artwork that carries neither an embedded PDF nor a bitmap is no longer dropped.
  Covers path brackets, the poly line/gon/bezier families, brush and pen objects, map modes,
  world transforms and DC save/restore; text, blits, clipping and arcs are skipped.

### Fixed

- Derive table height from the larger of the graphic frame extent and the row-height sum. `a:tr@h`
  is a minimum row height that PowerPoint grows to fit content, so deriving height from the grid
  alone crushed real-world tables by 15-50% and clipped cell text.
- Wrap an implicit single-line label instead of squeezing it onto one line when the shape omits
  `bodyPr@wrap`. OOXML defaults that attribute to `square`, so such a label wraps; squeezing is
  now limited to boxes where a single unwrapped line does not even fit the height.
- Stop rendering PPTX-embedded fonts by default. PowerPoint subsets embedded faces to the glyphs it
  believes are used, and a subset missing a glyph makes the browser fall back per character, so a
  single run rendered in two typefaces at two apparent sizes.

### Changed

- The python-pptx corpus generator now supports native PDF export on macOS, repeatable exact/glob
  case filters, and SHA-256 artifact records while keeping cached case metadata synchronized.
- Map Office percentage line and paragraph spacing through its native line unit, trim outer
  first/last paragraph spacing, and use text-container defaults validated by the 12-case CJK matrix.
- Stage macOS PowerPoint input/output in one fixed ignored runtime directory and use a bounded
  timeout so local corpus generation does not require a new folder grant for every case.
- Make the native macro smoke validate a non-empty SmartArt catalog produced in the fixed runtime
  directory.
- Install `pytest-timeout` for the declared 180-second E2E limit and remove the unused
  `asyncio_mode` setting, so pytest no longer ignores both configuration keys with warnings.
- Allow the resolved `pdfjs-dist` root in the Vite development server so PDF Worker browser tests
  remain valid when a Git worktree resolves dependencies outside its own directory.
- Honor `PPTX_E2E_BROWSER_CHANNEL` in pytest browser fixtures as well as the evaluation API, so
  local E2E runs can consistently use an installed branded Chrome.

### Fixed

- Keep near-fit, single-paragraph square-wrapped headings on one line with a conservative 2%
  browser-metric correction while preserving deliberate multi-line text.
- Resolve macOS PowerPoint exports and macro hosts by exact full path, close only that presentation,
  and qualify VBA procedures with the host filename so unrelated open decks cannot become the
  export/close target and unqualified macros do not fail with `-18`.
- Preserve AppleScript stderr, including environment codes such as `-9074`; bound exports and
  macros, remove stale output before native execution, and classify timeouts as an unlocked-session
  or pending permission problem instead of retrying a blocked UI.
- Generate CJK soft line breaks as DrawingML `a:br` elements so local oracle inputs do not contain
  visible `_x000B_` escape text.

- Align default vertical column, line, area, scatter, and bubble plot areas and side legends
  more closely with PowerPoint while preserving manual layouts, overlay legends, negative-value
  columns, and horizontal bars.
- Wait for fonts, images, and stable chart canvas output before oracle screenshots without
  changing ECharts animation or the public `SlideHandle.ready` contract.
- Select one compatible MCE Choice or Fallback across slide/template/group content and
  OLE picture previews, including the supported SVG picture extension and lazy paths.
- Preserve slide/layout/master color-map identity/reset semantics and isolate chart-local
  overrides from the parent slide.
- Follow matched layout placeholder categories during master inheritance, preserve explicit
  zero transforms/insets, and resolve body-property/autofit choices and text-container whitespace.
- Preserve sparse scatter/bubble coordinates and literal chart data; respect explicit negative-bar
  inversion flags and merged/conditional table borders, including corner styles and no-fill clears.
- Keep clipped picture effects and asynchronous media work attached to their owning render
  handle; preserve external handles across viewer reload/destroy and cancel disposed chart setup.
- Restrict segmented-cycle geometry compensation to matching SmartArt layout provenance.

- Prevented tolerated text metric overhang from turning PowerPoint text boxes into browser
  scroll containers, which could expose scrollbars and change wrapping on Windows.
- Sized tables from their column/row grid (Σ column widths × Σ row heights) instead of the
  graphicFrame `<a:ext>`, so tables authored in Google Slides — which export a stale
  placeholder ext — no longer render squished with clipped cell text.
- Fixed group child coordinate remapping for flat groups whose child extent is zero on one
  axis (e.g. a divider/underline built from horizontal connectors). The populated axis is now
  offset and scaled correctly instead of skipping the remap, which had left the children
  displaced by the group's child offset.

- Give every CJK font family a fallback chain. Only eight hardcoded names were recognised, so a
  deck asking for an uninstalled foundry font emitted a bare `font-family` and the browser fell
  back to its default standard font — a Latin serif on Windows, tofu without CJK coverage.
- Index media stored outside `ppt/media/`. OPC permits a picture next to the part that references
  it, and those parts were dropped instead of rendered; they now also count toward the media budget.
- Paint chart series picture fills (`a:blipFill`) instead of falling back to the palette, sized and
  anchored to the bar they paint for both the `stretch` and `stack` picture formats. Picture-filled
  bubbles use an image symbol, and legend swatches show the artwork rather than a palette colour.
- Keep a bottom chart legend clear of the value-axis labels by measuring the legend overlay instead
  of reserving a constant that only held at one text size.
- Place pie and doughnut charts from the `c:plotArea` manual layout, and leave a data point with
  `a:noFill` transparent so it exposes what sits behind the chart.
- Honour the blipFill source crop, tiling and geometry clipping on EMF pictures, which previously
  bypassed the shared image path and stretched a cropped band to its full frame.
- Shadow picture-filled shapes along their geometry. The effect fell back to a wrapper box-shadow
  that traced the bounding rectangle, drawing a straight line across a rotated custom geometry.
- Keep an unanchored single-line autofit text box top-aligned. `a:pPr@algn` is horizontal
  alignment and `a:bodyPr@anchor` defaults to top, so centring conflated the two and pushed text
  down onto whatever sat below it.

## [1.2.4] - 2026-07-10

### Added

- Added a standalone `./browser` ESM entry and real Chromium package tests covering PPTX
  rendering, all supported ECharts series, text overflow combinations, and PDF.js 5/6
  Worker compatibility.

### Changed

- ECharts now uses modular `echarts/core` registration, reducing the standalone browser
  bundle while preserving the renderer's chart support matrix and existing behavior.
- Production builds now use a cross-platform Node build script.
- PDF.js remains optional and external; no-bundler integrations should use pinned module
  and worker URLs and allow blob Workers in their CSP.

### Fixed

- Fixed compact no-wrap text runs creating horizontal scrollbars when their rendered
  width exceeded narrow PowerPoint text boxes.
- Fixed mixed horizontal/vertical overflow settings being converted by browsers into an
  unintended scroll container.
- Fixed isolated PDF.js cleanup on loading failure and guaranteed Worker termination on
  success, error, timeout, or cancellation.
- Fixed late EMF-PDF results mutating disposed slide DOM or repopulating shared blob URL
  caches after `SlideHandle.dispose()` or `PptxViewer.destroy()`.
- Replaced unsafe backtracking parsers for untrusted CSS and SVG path values with bounded
  parsing paths.

### Migration Notes

- No API migration is required. Applications with a restrictive CSP that enable
  EMF-PDF fallback must allow the configured PDF.js module source and `blob:` Workers.

## [1.2.3] - 2026-07-01

### Fixed

- Fixed connector arrowhead direction, sizing, and placement for flipped or transformed
  connector paths so arrow markers better match PowerPoint output.
- Fixed compact numeric text such as `80%` and adjacent numeric/unit runs so browser
  wrapping no longer splits the number from its percent or unit marker.

### Migration Notes

- No migration is required.

## [1.2.2] - 2026-06-28

### Added

- Improved the E2E comparison review UI so manual PDF-vs-HTML inspection has clearer
  metrics and review state handling.

### Fixed

- Improved chart fidelity across axis density, tick visibility, label sizing, legend
  order, legend margins, marker defaults, blank data points, data table semantics, rich
  chart titles, text shadows, plot-area backgrounds, and radar manual layout.
- Improved pie, doughnut, radar, scatter, bubble, horizontal bar, stacked, stock, and
  dense line chart defaults so compact and Office-authored charts render closer to
  PowerPoint.
- Fixed gradient rendering for path and background fills, including themed path
  gradients, focus rectangles, pixel-space radial radii, and subpixel gradient strokes.
- Fixed picture and shape effects including clipped picture fills, grayscale picture
  effects, inner shadows, soft edges, and scaled-down shape shadows.
- Fixed table and text edge cases for outer paragraph spacing, paragraph tab defaults,
  vertical WordArt, and Office-like font fallback behavior.
- Fixed shape and SmartArt edge cases including fillable brace presets, flowchart
  storage guide alignment, and SmartArt cycle pie offsets.

### Migration Notes

- No migration is required.

## [1.2.1] - 2026-06-25

### Fixed

- Fixed grouped content with horizontal and vertical flips so child geometry, readable
  text behavior, and connector arrow directions better match PowerPoint output.
- Fixed flipped picture rendering when shape clipping and image crop metadata are both
  present, preserving the image orientation inside the flipped clip shape.
- Fixed table and chart frames inside flipped groups so their positions mirror with the
  group while their table/chart content remains correctly oriented.
- Improved compact column chart rendering by reducing over-dense default value-axis
  ticks for small chart frames.

## [1.2.0] - 2026-06-16

### Added

- Added lazy media decoding via `parseZipLazyMedia()` and `lazyMedia` options on
  `PptxViewer` and deprecated `PptxRenderer`.
- Added lazy slide node parsing via `buildPresentation(files, { lazySlides: true })`,
  `materializeSlideNodes()`, and `materializeAllSlideNodes()`.
- Added exported lazy-materialization helpers for search, serialization, and custom
  model consumers that need explicit control over deferred slide nodes.
- Added a repeatable performance benchmark tool under `test/perf/render_benchmark.py`
  for comparing eager, lazy, full-list, and windowed rendering paths.
- Added public demo coverage for recommended large-deck options.

### Changed

- Public demo rendering now uses recommended ZIP limits, lazy slide parsing, lazy media
  decoding, and windowed list mounting by default.
- Public demo empty state now exposes direct upload and sample actions.
- Public demo no longer exposes the broken list/slide mode toggle; it focuses on the
  stable scrollable list workflow with search and thumbnails.
- README now documents when to use windowed rendering, `lazySlides`, `lazyMedia`, and
  eager rendering for export-style workflows.
- Performance documentation now covers lazy media, lazy slide parsing, benchmark
  expectations, and strategy selection for small, medium, and large decks.
- Chart rendering internals were split into smaller modules while preserving the public
  API.

### Fixed

- Improved rendering fidelity across OOXML chart defaults, legends, data labels, axes,
  table styling, text autofit, arrow markers, pattern fills, shadows, and template
  preview content.
- Hardened OOXML parsing for relationship targets, namespaced attributes, package media
  path aliases, and boolean default semantics.
- Improved serialization and search compatibility with lazily materialized slide nodes.

## [1.1.0] - 2026-06-05

### Added

- Added model-level text search APIs: `buildTextIndex()`, `searchText()`,
  `searchPresentation()`, and `PptxViewer.searchText()`.
- Added exported search result and option types, including `TextSearchResult`,
  `TextIndexEntry`, `TextSearchOptions`, `TextIndexOptions`, and `TextBounds`.
- Added node-level search highlight helpers on `PptxViewer`:
  `highlightSearchResult()` and `clearSearchHighlights()`.
- Added `SearchHighlightOptions` so consumers can customize highlight class names,
  colors, border width, radius, padding, shadows, z-index, and inline styles.
- Added scaled slide preview rendering via `renderThumbnailToContainer()` for
  thumbnail/navigation surfaces.
- Added lazy thumbnail navigation and search result navigation to the dev page and
  public demo.

### Changed

- String text searches are case-insensitive by default and can opt into exact casing
  with `matchCase: true`.
- RegExp text searches now preserve caller-provided flags; the search layer only adds
  `g` so all matches can be collected.
- Dev/demo thumbnail cards are smaller and keep active selection styling layout-stable.
- Search UI in dev/demo now includes an `Aa` match-case toggle and uses the viewer
  highlight API instead of rewriting rendered text.
- Documentation now covers search/highlight API boundaries, thumbnail performance
  trade-offs, and the distinction between scaled DOM/SVG previews and bitmap
  thumbnail generation.

## [1.0.4] - 2026-06-02

### Added

- Added chart regression coverage for 100% stacked bars, stacked line/area charts, reversed axes, per-point pie labels, multi-ring doughnut charts, and chart color-style palettes.
- Added `--testdata-source=windows|all` for pytest E2E runs so Windows-generated chart oracle cases can be exercised without changing test code.
- Added configurable PDF.js fallback support via `pdfjs` options on `PptxViewer`, legacy `PptxRenderer`, and headless `renderSlide()`.
- Added exported `PdfjsOptions` and `PdfjsConfig` types for consumers that need EMF-embedded PDF preview rendering.
- Added `SlideHandle.ready` coverage for async slide resources such as EMF-PDF fallback previews.
- Added repository `AGENTS.md` guidance for future agent-assisted maintenance.

### Changed

- Chart rendering now honors OOXML stacked and 100% stacked grouping for bar, line, and area charts, including percentage axis scaling.
- Chart rendering now applies closer Office defaults for bar `gapWidth`, 100% stacked tick intervals, and area-chart value-axis headroom.
- Doughnut charts with multiple series now render as concentric rings instead of dropping later series.
- Implicit chart palettes now use chart color-style parts when related from the chart part.
- Documentation now distinguishes chart 3D graceful fallbacks from true 3D chart fidelity.
- PDF.js is no longer bundled into the core library output by default; consumers can provide `moduleUrl` and `workerUrl` only when they need EMF-PDF fallback rendering.
- The local test and review pages now configure PDF.js explicitly for EMF-PDF preview rendering.
- Documentation now clearly explains that ordinary PPTX rendering does not require PDF.js configuration and that full EMF/WMF vector rendering remains out of scope.
- Rendering fidelity was refined across Office text layout, picture shape properties, chart line styles, table overrides, explicit strokes, theme fonts, and async media readiness.

### Fixed

- Fixed reversed chart axes (`orientation="maxMin"`) being parsed but not applied.
- Fixed chart axis lookup falling back to the first axis before checking all matching `axId` values.
- Fixed pie charts ignoring per-point `c:dLbl` position, style, leader-line, and manual-layout overrides.
- Fixed line, scatter, and combo chart cases that could drop later plot-area series or lose OOXML smoothing/line semantics.
- Fixed picture shape properties such as fills, outlines, shadows, and related rendering details being ignored for pictures.
- Fixed text sizing, wrapping, bullet colors, vertical labels, hyperlink colors, and theme-font inheritance mismatches seen in real-world decks.
- Fixed EMF-embedded PDF fallback icons not appearing in screenshot/export flows when PDF.js worker URLs were not isolated correctly.
- Fixed async EMF-PDF fallback rendering so callers can await `SlideHandle.ready` before visual capture.
- Fixed explicit shape lines without a width so Office's visible default stroke width is preserved.
- Fixed chart legend, tooltip, data-label, and axis label sizing/positioning regressions across Oracle and real-world decks.

## [1.0.3] - 2026-05-26

### Added

- **Recommended ZIP safety limits** — `RECOMMENDED_ZIP_LIMITS` provides documented defaults for rendering untrusted PPTX input.
- **Expanded python-pptx placeholder oracle coverage** — added a regression case for idx-only placeholder inheritance, covering layout/master text style resolution and bullet sizing.
- **E2E comparison slide URL state** — the selected slide is now reflected in the `slide` query parameter so manual visual review survives refresh and shared URLs.
- **Chart lifecycle regression coverage** — standalone `renderSlide()` chart instances are now covered to ensure charts are disposed when the returned slide handle is disposed.
- **Rendering fidelity unit coverage** — added regression coverage for real-world large-deck rendering issues.

### Changed

- **Resource limits are enforced against decoded entry sizes** when ZIP metadata is unavailable, improving protection against malformed or adversarial PPTX archives.
- **Render queue cancellation is stricter**: stale batched list renders stop when a newer render request supersedes queued work.
- **External media handling is safer by default**: unsafe external media relationships are rejected and media preloading is disabled during rendering.
- **Chart rendering is closer to Office output**: chart-local theme overrides, combo chart series, radar legend layout, default chart typography, axis label sizing, data labels, and interactive label sizing now follow OOXML semantics more closely.
- **Text layout now honors more Office body properties** including inherited `bodyPr`, autofit modes, hyperlink theme colors, inherited bullet colors, arched text transforms, and narrow CJK vertical-style labels.
- **Shape and layout rendering defaults were refined** across fills, strokes, groups, tables, images, and backgrounds to reduce browser-default drift from PowerPoint output.

### Fixed

- Fixed resource exhaustion risks from oversized decoded ZIP entries, chart cache point allocation, and EMF bitmap decoding.
- Fixed stale batched renders mutating the DOM after a newer render cycle had started.
- Fixed placeholder-only slide shapes losing layout/master text style inheritance, including title/body categories and inherited bullet sizing.
- Fixed bullet glyphs rendering too small when font size came from inherited `defRPr` plus `normAutofit` scaling.
- Fixed chart-local theme overrides being ignored in some cases.
- Fixed combo charts dropping later plot-area series, such as additional line series.
- Fixed radar chart legend positioning and sizing issues.
- Fixed chart legend labels, data labels, axis labels, and hover labels using incorrect default sizes.
- Fixed hyperlink text color inheritance for links without explicit run colors.
- Fixed inherited bullet colors rendering incorrectly on dark backgrounds.
- Fixed arched text effects being flattened to ordinary straight text.
- Fixed text autofit and body positioning mismatches that could cause overflow, unexpected shrinkage, or shifted text relative to its shape.
- Fixed narrow CJK chart labels being shrunk instead of rendered as wrapped vertical text.
- Fixed table, image, background, and group rendering defaults that diverged from Office in real-world decks.
- Fixed standalone `renderSlide()` chart lifecycle cleanup while preserving caller-owned chart instances.

## [1.0.2] - 2026-03-09

### Added

- **Python-pptx ground truth pipeline** — second test case generation pipeline using `python-pptx` for PPTX creation and PowerPoint COM for PDF/PNG export. Generates 100 new cases (`oracle-pypptx-*` prefix) covering rich text (38 cases), shape adjustments (31 cases), composite layouts (10 cases), and chart variants (21 cases).
- **Expanded VBA ground truth catalogs** — fill/stroke configs from 10 to 36 (new solids, gradients, patterns, dash styles, colored strokes), table configs from 7 to 15 (edge cases like 1×1, 10×1), connector configs from 6 to 9 (remaining orientations), and dynamic chart type probe with 103-entry `XlChartType` fallback dict.
- **Visual regression cases**: 352 → 452+ total automated cases, all passing with zero failures.
- **Unit tests**: 1400+ new lines of test coverage for ChartRenderer (lifecycle + rendering), ShapeRenderer, StyleResolver, TableRenderer, and preset shapes.

### Changed

- **Chart color fidelity** — use theme accent palette (`option.color`) instead of hardcoded `DEFAULT_SERIES_COLORS` for scatter, bubble, and radar series fallback colors. Add candlestick up/down colors from OOXML series `spPr` in stock charts.
- **Scatter/bubble axis handling** — new `parseScatterAxes()` correctly parses two `valAx` nodes by axis position (`b/t` → X, `l/r` → Y). Fix gridlines direction: Y-axis `majorGridlines` now render as horizontal lines.
- **Scatter chart markers** — parse `scatterStyle` (`lineMarker`/`smoothMarker`) to default diamond markers. Apply OOXML marker symbols and sizes per series.
- **Auto-title for all chart types** — pass `seriesArr` to `extractChartTitle()` in all builders (bar, line, scatter, bubble, radar, stock) enabling auto-generated title from series name when `autoTitleDeleted=0`.
- **Legend icons** — respect per-item OOXML marker symbols (circle, diamond, triangle) for line/area/radar instead of always overriding to `rect`.
- **Radar chart** — add `areaStyle` with semi-transparent fill (0.15 standard, 0.5 for filled style).
- **Bar width formula** — fix `barCategoryGap` to account for number of series in clustered bars: `gapWidth / (100 × N + gapWidth)` instead of `gapWidth / (100 + gapWidth)`.
- **Chart font size propagation** — extend `applyDefaultFontSizes` to also override series data label font sizes with `chartSpace txPr` default when no explicit OOXML font was set.
- **Windows ground truth pipeline** — per-case independent COM sessions for fault isolation, retry helpers for Windows file handle races, absolute output paths for VBA `SaveAs`.

### Fixed

- Scatter/bubble/radar series now use theme accent colors consistently instead of hardcoded palette.
- Stock chart candlestick colors now read from OOXML series style properties.
- Clustered bar chart width calculation now correct for multi-series charts.
- Chart data labels inherit `chartSpace txPr` default font size when no explicit size is set.
- Windows COM ground truth generation: `RPC_E_CALL_REJECTED` retry with exponential backoff, `pres.Saved=True` before `Close()` to suppress save dialogs.

## [1.0.1] - 2026-03-01

### Added

- **`PptxViewer`** — new recommended API class extending `EventTarget`. Separates parsing, model loading, and rendering into distinct steps:
  - `PptxViewer.open(input, container, options?)` — static factory that parses, builds, and renders in one call.
  - `viewer.load(presentation)` — load a `PresentationData` model without rendering.
  - `viewer.renderList(options?)` — render all slides in a scrollable list.
  - `viewer.renderSlide(index?)` — render a single slide (no built-in nav UI).
- **`SlideHandle`** — per-slide resource lifecycle returned by `renderSlide()` and `renderSlideToContainer()`. Tracks chart instances and blob URLs for deterministic cleanup via `handle.dispose()`.
- **`ListRenderOptions`** — dedicated options type for `renderList()`: `windowed`, `batchSize`, `initialSlides`, `overscanViewport`.
- **EventTarget events** — `slidechange`, `sliderendered`, `slideerror`, `slideunmounted`, `nodeerror`. Typed via `PptxViewerEventMap`. Shorthand callbacks (`onSlideChange`, etc.) also supported.
- **`Symbol.dispose`** — `PptxViewer` implements TC39 Explicit Resource Management (`using viewer = ...`).
- `scrollContainer` option: custom scroll root for `IntersectionObserver` in windowed list mode.
- `onSlideUnmounted` callback / `slideunmounted` event: fires after a slide is unmounted in windowed list mode.
- `isSlideMounted(index)` and `getMountedSlides()` methods: query which slides are currently mounted in the DOM.
- `AbortSignal` support in `PptxViewer.open()` and `PptxRenderer.preview()`.
- `ScrollIntoViewOptions` parameter in `goToSlide(index, scrollOptions?)`.
- Scroll-based slide tracking in list mode via `IntersectionObserver` (fires `slidechange` for the most-visible slide).
- **`renderstart` / `rendercomplete` events** — bracket every render cycle (renderList, renderSlide, setZoom, setFitMode). `rendercomplete` fires even when render throws.
- **`isRendering` getter** — `true` between `renderstart` and `rendercomplete`.
- **`on()` / `off()` typed event helpers** — convenience wrappers over `addEventListener`/`removeEventListener` with proper generics. Returns `this` for chaining.
- **`zoomPercent` / `fitMode` getters** — read current zoom level and fit mode.
- **Instance-level `open()` method** — parse, build, and render from binary input on an existing viewer. Cleans up previous state on re-open. Static `PptxViewer.open()` now delegates to this.
- `onRenderStart` / `onRenderComplete` shorthand options in `ViewerOptions`.

### Changed

- `renderSlide()` (from `SlideRenderer`) now returns `SlideHandle` instead of `HTMLElement`.
- `renderSlideToContainer()` now returns `SlideHandle` instead of `HTMLElement | null`.
- `onSlideChange` now fires in both list mode (scroll tracking) and slide mode (navigation). Previously only documented for slide mode.
- **`slidechange` now fires after every render cycle** (renderList, renderSlide, setZoom, setFitMode), reporting the current slide index. This means consumers always receive an initial `slidechange` after the first render.
- **`goToSlide()` now returns `Promise<void>`** instead of `void`. In list mode, resolves after initiating mount + scroll. In slide mode, resolves synchronously.
- **`renderSingleSlide` error handling** — errors in slide mode now show an error placeholder (consistent with list mode) instead of propagating.
- `pdfjs-dist` moved from `dependencies` to optional `peerDependencies`. Install separately if using SmartArt PDF fallback rendering: `npm install pdfjs-dist`.

### Deprecated

- **`PptxRenderer`** — use `PptxViewer` instead. `PptxRenderer` extends `PptxViewer` and provides the legacy `preview()` API with built-in nav buttons in slide mode.
- **`RendererOptions`** — use `ViewerOptions` instead.

### Fixed

- `renderSlideToContainer()` now passes `chartInstances` to `renderSlide()`, preventing ECharts memory leaks in external containers.
- `renderSingleSlide()` (slide mode) now passes `chartInstances` to `renderSlide()` for proper chart lifecycle tracking.
- Main-thread pdfjs fallback no longer sets `GlobalWorkerOptions.workerSrc` to a URL, eliminating global pollution when host apps use their own pdfjs instance.

## [1.0.0] - 2026-02-28

### Added

- Browser-side PPTX parsing and rendering (`list` and `slide` modes).
- **Shape geometry**: 187+ preset shapes from ECMA-376 spec, plus custom geometry (`<a:custGeom>`) interpreter. 33+ multi-path 3D shapes with lighten/darken face modifiers.
- **Text rendering**: 7-level OOXML style inheritance, theme fonts, numbered/symbol/picture bullets, vertical text, superscript/subscript, hyperlinks.
- **Charts**: bar, line, area, pie, doughnut, radar, scatter, surface (2D and 3D variants) via ECharts.
- **Fill & stroke**: solid, linear/radial/rectangular gradient, 52+ pattern fills, image fills; 8 dash styles, 5 arrowhead types, compound lines.
- **Color pipeline**: full OOXML resolution — schemeClr → colorMap → theme lookup → modifiers (lumMod, lumOff, tint, shade, alpha, satMod, etc.). All 6 color spaces supported.
- **SmartArt**: 134+ layouts via PowerPoint fallback data.
- **Tables**: OOXML table styles, cell merge (gridSpan + rowSpan), border inheritance.
- **Images**: blob URL rendering with crop, stretch/tile, video/audio placeholders.
- **Groups**: coordinate remapping (chOff/chExt) with recursive child rendering.
- **Backgrounds**: slide → layout → master inheritance chain (solid, gradient, image, pattern).
- **Security**: ZIP parsing limits (`ZipParseLimits`), external hyperlink protocol filtering.
- **Performance**: windowed list mounting via `IntersectionObserver`, batch rendering, large-deck tuning knobs.
- **Visual regression testing**: 352 automated cases (187+ shapes, 134+ SmartArt, 37 fill/stroke variants) verified against PowerPoint output using SSIM + color histogram correlation. Zero failures.
- **Quality tooling**: ESLint, Prettier, commitlint (Conventional Commits), husky pre-commit hooks, knip (dead code detection), publint, size-limit.
- **Documentation**: architecture, testing, performance, contributing, security, and releasing guides.

### API

- Main class: `new PptxRenderer(container, options)`
- Core render call: `await renderer.preview(input)` where `input` is `ArrayBuffer | Uint8Array | Blob`
- Navigation/lifecycle: `goToSlide(index)`, `destroy()`
- Runtime scaling: `setZoom(percent)`, `setFitMode('contain' | 'none')`
- Utility exports: `parseZip`, `buildPresentation`, `serializePresentation`
