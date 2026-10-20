# Validation

## Cheat Sheet rollout refresh — 20 October 2026 — app version 1.0.21.0

Adopted the user's manually edited `Document Comparison Tool Cheat Sheet.docx` as the maintained user handout. The document is copied unchanged into the team package and installation. Git explicitly includes the new filename and removes the former QuickGuide. The template generator now writes an ignored draft only, preventing it from overwriting manual edits. Builds require the maintained file and remove the old guide from the release folder; installation replaces the old package-managed guide after copying the new document.

The source document passed ZIP/relationship integrity checks and contains no external links, macros, comments or tracked changes. Release verification compares its exact SHA256 (28B51D014C03B323FF1F535A96CC726A96CCB1C517E4ABF48903B47F7D627325) between the source, ZIP, installed copy and desktop handout. The 30-file package and checksum, extracted app startup and installer migration are checked; saved results and an existing pilot log are preserved. This documentation/packaging change leaves comparison behavior and app version 1.0.21.0 unchanged, so the earlier functional checks remain applicable.

## Team rollout refresh — 15 October 2026 — version 1.0.21.0

The user confirmed that the local OCR layout update works and approved updating the team package and public repository. Version 1.0.21.0 includes the tested OCR geometry helper and Word layout builder, with the current RWS quick guide, acceptance checklist and updated Word/text test plan. Recognition, counting, crossover and language-check rules are unchanged. The documented approximation of fonts, complex columns, tables and graphics still applies.

Release checks verify all 30 intended ZIP files and their hashes, the checksum sidecar, direct startup of the extracted app, the three tabs, the current guide and empty pilot-log template. The extracted installer preserves saved results and the existing installed pilot log. The packaged OCR path is checked with synthetic scanned PDFs for accurate counts, yellow insertions, complete pages and retained page sizes, heading sizes, spacing and indents. The previous 1.0.20 team ZIP is retained privately for recovery; the refreshed package and guide are copied to the desktop for sharing. The source change is committed and pushed to the existing public repository after those checks.

## Version 1.0.21 — 15 October 2026 — local OCR-layout trial; distribution held

The OCR path now retains numeric page and line geometry alongside its unchanged canonical recognised text. A new Word builder reconstructs source page sizes/orientation, approximate text sizes, indents and paragraph spacing, retaining one paragraph per recognised line and the established reading order. Mixed page sizes and sparse first-line placement are supported. Missing or invalid geometry uses the previous plain-document path with a review note. No font-family recognition, semantic table recovery, scan-image overlay or column reordering is claimed; complex columns can flow sequentially. Specialist review against the source remains necessary, and a checked ABBYY DOCX is the alternative when close reproduction is required.

Passed 123 focused text/layout checks, including identical word counts, insertion counts, text fingerprints and final-text order before/after formatting; source page geometry; headings/indents/gaps; mixed sizes; sparse captions; invalid/missing metadata fallback; and no overwrites. Existing suites passed 55 counting/revision/highlight checks, 79 EP checks, 37 pilot interface checks and 74 Multi Archive checks. The existing Windows OCR smoke test passed mixed selectable/scanned pages, text/blank pages, entirely blank refusal, conflicting paths and invalid PDFs without source modification or document-text disclosure.

Passed 34 actual Windows OCR/desktop Word checks on two-page synthetic scanned PDFs. Both old and new paths returned 105 words with one insertion and 99.047619% crossover for the changed pair, and 105 words with zero insertions and 100% for the identical pair. The saved comparison changed from the old A4 paper size to the original 612 by 792 points, retained distinct heading/body sizes (23.5/12.5 points), a 93-point visible paragraph gap instead of 51, zero space after instead of Word's default 8 points, and a 277.2-point right-column paragraph indent instead of zero. Known source words appeared once, canonical OCR line order was preserved, insertions remained yellow and source hashes were unchanged. Private Word sessions closed; an optional PDF-preview export stalled and was stopped by an identity-checked watchdog, without affecting the saved DOCX or comparison results.

The app and OCR helper are installed locally for user testing. The team ZIP/released folder remain at 1.0.20.0 with SHA256 0717228B842869AEA69E02A0B84AED1DE1AED6D1380F933ECB543A7620119AE9, and no GitHub commit or push is part of this trial. Saved results and the installed pilot log are preserved. The local RWS guide reflects the trial; the distributed guide remains unchanged.

## Team rollout refresh — 12 October 2026 — version 1.0.20.0

Refreshed the team distribution from 1.0.17.0 to the current 1.0.20.0 app, including the EP data-table/whitespace fixes, Multi Archive and archive-first layout. The package includes the current three-page RWS quick guide, current acceptance checklist and refreshed Word/text test plan. The release test template now has 44 scenarios and 63 fresh Not run rows, including Multi Archive selection/ranking, EP ranking by Full specification, language/OCR checks, cancellation/shared reports, EP preparation regressions and Reset. No team test outcomes have been marked as passed; existing installed pilot logs are retained during upgrades.

The rebuilt release passed actual desktop Word checks using synthetic files: Full Document insertion comparison returned 5 new words, 1 inserted word and 80.0% crossover; identical prepared EP inputs returned Description 9 words, Full specification 21 words and Claims 12 words, each with zero insertions and 100.0% crossover. Source hashes were preserved and both private Word processes closed. Earlier version-specific functional checks below remain applicable; this refresh changes distribution and documentation rather than comparison logic.

Distribution verification checks the 30-file ZIP inventory and every file against the release folder, the SHA256 sidecar, extracted-app startup with all three tabs, the guide version and uncompleted log template, and a per-user update using the extracted installer while preserving saved results and the existing pilot log. The desktop package and separate quick guide are copies of the verified release. Running DocumentCompare.exe opens the app directly; installation and manual upgrades use Install.cmd after the complete ZIP is extracted. A private copy of the previous team ZIP is retained for recovery.

## Version 1.0.20 — 9 October 2026 — local update; team package deferred

Moved the Multi Archive folder picker, archive list and archive actions above New document, followed by the results destination. Keyboard navigation follows that order, Reset focuses the archive folder field and the on-screen instructions and RWS quick guide describe archives first. Comparison, ranking and file-loading logic are unchanged.

Passed the existing 54 Multi Archive interface checks and 37 pilot-feedback interface checks, including selection, OCR binding, tab locking, Reset and widths at normal and minimum window sizes. Rendered and inspected both sizes; the normal window shows archives, New document and save controls in order, while the minimum window scrolls. Private Word repagination confirmed the updated quick guide remains three pages. The desktop app and installed guide are updated locally; the team ZIP and released folder remain unchanged at 1.0.17.0.

## Version 1.0.19 — 9 October 2026 — local update; team package deferred

Added **Multi Archive**, a third tab comparing one fixed new document against each selected archive. Full Document ranks by whole-document crossover; EP ranks by Full specification crossover and retains all three scope rows. Archives can be added with a multiple-file picker, drag/drop or folder loading; folder loading replaces the archive list and excludes the selected new file. One archive is permitted. Source/OCR selections are snapshotted into a dedicated one-to-many plan, with unique output folders and `multi-archive-results.csv`.

The interface and CSV order verified scores highest first, share ranks for exact ties and leave unavailable, unsuccessful, review-pending and unfinished results unranked. Counts alone cannot restore a withheld percentage. Running ranks are labelled provisional; the final summary names the best eligible archive(s), and cancellation explicitly limits that summary to completed comparisons. Existing worker isolation, local processing, language checks, report retry, cancellation and cleanup are shared with Batch comparison. All three tabs prevent concurrent jobs. Multi Archive Reset also clears its fixed new file/OCR and returns to the top of the form while preserving mode, destination and saved outputs. Completion brings the results into view.

Passed 74 Multi Archive model/ranking/CSV checks and 54 actual WinForms checks for selection, duplicates, OCR binding, folder loading, ranking, minimum-width layout, tab locking and Reset. Passed 67 actual hidden WinForms/desktop Word checks using synthetic documents: one new file against four archives, a known language mismatch followed by successful matches and one owned summary dialog, independently expected insertion counts and crossover, actual yellow markup, sorted CSV, EP ranking by Full specification when Description favours another archive, all three EP scopes, translated-claims exclusion, uncertain-language pause/current-pair retry, and cancellation retaining completed outputs and partial rankings. Source hashes were preserved and private jobs closed. The final best-match presentation/Reset scroll and Skip archive wording passed the focused UI checks after those final adjustments.

Existing suites passed 49 batch model/report, 37 batch UI state, 156 batch Reset, 106 single Reset, 37 pilot-feedback UI, 79 EP section and 30 native language-dialog checks. The RWS guide now includes selection and operating guidance for Multi Archive; private Word repagination confirms a three-page handout. README and the manual acceptance checklist include the new mode. The team ZIP and released folder remain unchanged at 1.0.17.0 as requested; this update is local only.

## Version 1.0.18 — 9 October 2026 — local update; team package deferred

Fixed a false EP preparation refusal when an otherwise well-structured specification includes a Description column in a data table. The exception is limited to a non-leading, single-paragraph column label in the first row of a top-level, unmerged data table with populated cells below, inside independently identified standalone Description and Claims boundaries. Nested claims, translations, exclusion headings, text boxes, content controls and ambiguous section layouts still stop preparation. The entire table remains in the description and its counts.

Also fixed normalization of non-breaking and narrow non-breaking spaces at Word run boundaries. Replacing them with ordinary spaces now sets `xml:space="preserve"` where needed. Previously Word could discard these spaces, join words and correctly trigger the final-text reconciliation safeguard. The safeguard and crossover formula are unchanged.

Passed 79 EP boundary, scope, revision, numbering, language-label and source-preservation checks, including data-table retention, conservative rejection cases and whitespace preservation. The provided real Word file reproduced the original nested-heading refusal. After both fixes, an actual private Word self-comparison returned Description 12,883 words, Full specification 13,497 words and Claims 614 words; every scope had zero insertions and 100.0% crossover, with strict final-text reconciliation passing. Local language detection identified English for Description and Claims. The original source hash stayed unchanged, the private Word session closed and all confidential diagnostic/working copies were removed. The team ZIP and released folder were deliberately left unchanged at the user's request.

## Version 1.0.17 — 8 October 2026

The RWS quick guide now starts with a selection table: Full Document by default, EP for EuroFile multi-language jobs needing separate Description, Full specification and Claims figures, and the Batch comparison tab for multiple files. Batch retains the same Full Document/EP selection rules. The regenerated document retains its RWS header and two-page layout, verified by repagination in a private Word session.

Added a clear single-comparison pop-up for a detected source-language mismatch, showing the compared scope and archive/new languages, explaining how to choose matching sources and confirming that no comparison document was created. The worker returns structured mismatch details; the UI does not infer the failure type by parsing the error text. Unknown/Mixed language still stops for the existing manual review, and manual confirmation cannot override a known mismatch.

Batch comparison retains failed-pair details and continues other pairs. One summary appears after the batch returns to idle, including after user cancellation when there are previously failed mismatch pairs. It lists at most five pairs with shortened filenames and directs users to the full table/CSV. Closing or disposing the app suppresses this summary, including delayed stop retries. Reset, report saving and repeated completion do not repeat it.

Passed 92 native Windows language-preflight/helper checks using local synthetic DOCX files: Full Document and EP Description/Claims mismatches, manual confirmation remaining blocked, Unknown/Mixed classification, source-byte preservation, no output or Word startup, JSON round trips and bounded summary formatting. Passed 30 actual native-dialog/UI checks for single and batch notices, cleanup before display, restored controls, Reset, cancellation/closing suppression and absence of duplicate notices after completion or report retry. The test harness uses bounded window reads and a 60-second watchdog; owned test processes and synthetic temporary files were cleaned up. Existing UI suites passed 37 pilot, 37 batch state and 156 batch Reset checks. The README, RWS cheat sheet and manual acceptance checklist now cover the notices.

## Version 1.0.16 — 7 October 2026

Added **Reset** beside Job type in both tabs. It clears selected files and PDF OCR choices, displayed figures/results, old language-review actions and batch session references so another job can start in the same window. It preserves the selected job type, results destination and all saved source/comparison/summary/report files. Active or paused jobs, attached workers and unconfirmed child shutdown prevent Reset. Batch Reset also invalidates queued callbacks, and either tab retains any leftover temporary-folder reference and warning.

Passed 106 single Reset checks, 156 batch Reset checks, 37 existing pilot UI checks and 37 batch UI state checks without starting Word. These cover completed/cancelled/failed work and manual language review, source/output byte preservation, new file selection, reset availability, retained cleanup notes, saved-settings preservation, stale callback rejection and placement at minimum window width. The updated single interface was rendered and inspected. The README, RWS cheat sheet and manual acceptance checklist now describe Reset. Word repagination separately confirmed the final guide remains two pages; its verified private check session was closed.

## Version 1.0.15 — 7 October 2026

Renamed the visible whole-document option to **Full Document** in both tabs and aligned the accompanying PDF guidance, README and user cheat sheet. The internal WF identifier and comparison calculations are unchanged. This wording update does not require further Word or shared-drive comparisons; the version 1.0.14 functional checks remain applicable.

## Version 1.0.14 — 7 October 2026

Added WF guidance for EP new documents supplied as PDFs or converted from PDFs. The application recommends whole-document comparison without changing the selected mode automatically; native Word EP section comparison remains available. The user guide explains preparing both inputs to the same intended scope and that WF returns one set of figures.

Made input Browse buttons prominent and fixed their off-screen placement at narrow window widths. Single file, batch folder and results-folder browsing remain disabled while their configuration is locked. Shared-drive sources and published results now accept UNC paths and mapped network drives, together with local and OneDrive paths. Conversion, OCR and Word comparison continue to use private local working copies outside known OneDrive locations; the UI validates that working path before creating it. Existing outputs are never overwritten.

Passed 40 storage-path checks, including two actual mapped drives, UNC inputs/outputs, OneDrive outputs, malformed/device paths and rejection of shared/OneDrive private work. Passed 55 metric/revision/highlight checks, 49 batch planning/report checks, 37 pilot-feedback UI checks and 37 batch UI state checks. Simulated OneDrive checks confirmed source preservation and the separate private-work restrictions. Both tabs were rendered and inspected. Batch result-button state no longer repeatedly queries a shared drive during grid refresh; opening an unavailable saved path still reports the error. CSV replacement retains its existing metadata/ACL preservation.

`tests/SharedDriveIntegration.ps1` passed 75 live shared-location checks in the approved team test folder. A single comparison read the archive and saved DOCX/TXT via UNC while reading the new source through the mapped drive: 5 new words, 1 yellow insertion and 80.0% crossover. The two aliases were verified to refer to the same files. The mapped-drive batch completed all six ordered pairs with expected counts/crossover, saved yellow comparisons, final CSV rows, locked-report preservation and retry, unchanged source hashes, and six cleaned local jobs. Both folder-load aliases passed, and cleanup refused paths outside owned local work. Only the unique synthetic files were removed; the approved Test folder was left empty. No native Word/helper test processes remained. Shared report replacement took several seconds per update on this network; the original sparse-log test was interrupted before completion, and the instrumented final rerun completed successfully. Interrupted synthetic artifacts were removed after ownership/process checks. Network-disconnection and read-only-destination scenarios remain pilot checks; the test did not alter server permissions or connectivity.

Updated the RWS Word cheat sheet, README, acceptance checklist, pilot plan and fresh pilot log. The pilot plan now has 38 scenarios; its 53 log rows remain Not run with blank actual figures. Word repagination verified the final cheat sheet remains two pages; PDF preview export timed out and its verified private preview process was closed by a bounded watchdog. Installation preserves an existing user's pilot log and saved results.

## Version 1.0.13 — 6 October 2026

Fixed false crossover withholding when Word's comparison contains moved-away paragraph marks (`w:pPr/w:rPr/w:moveFrom`). Like a deleted paragraph mark, a moved-away mark joins the surviving text to the following paragraph. The reader now preserves that final paragraph structure and insertion flags; the original word-count and text-fingerprint reconciliation checks remain unchanged.

`tests/EngineTests.cs` passed 55 checks, including 24 new regressions for moved paragraph joins, mixed deletion/move sequences, inserted word fragments, surviving move-destination marks, historical property snapshots, cell/text-box boundaries, language extraction and retained tracked markup.

All 15 comparisons in the reported local batch passed the original strict text/word checks with the corrected reader. The 14 previously withheld percentages were caused by moved paragraph marks; the total and insertion counts stayed unchanged. A fresh Word comparison of the first reported pair also passed: 10,343 new words, 7,020 insertions and 32.1% crossover. Original hashes stayed unchanged, the private Word session closed and confidential temporary working copies were removed.

A separate corrected batch copy regenerates the CSV and TXT statistics while preserving byte-for-byte identical Word comparisons. Every original source and previous result file remains unchanged. The pilot instructions now include moved paragraph breaks and crossover availability across all completed batch pairs; pilot results and synthetic fixtures were preserved.

Focused native Word regressions also passed: the established WF move fixture retained 90 words, zero insertions and 100.0% crossover; the combined EP edit fixture retained Description 10/2/80.0%, Full specification 27/8/70.4% and Claims 17/6/64.7%. Fixture hashes stayed unchanged and the private Word sessions closed.

## Version 1.0.12 — 6 October 2026

Added user instructions and manual pilot cases for folder-based Batch comparison: top-level DOC/DOCX/PDF selection, visible natural ordering and reordering, all later-versus-earlier pairs, per-PDF OCR, WF/EP report scopes, unique batch/pair output folders, language-review pauses, failed-pair handling and preservation of completed outputs after cancellation. The generated report is `batch-results.csv`; it is separate from the editable `PilotResults.csv` pilot log.

`tests/BatchPlanTests.cs` passed 49 planning/report checks: exact six-pair sequence, natural filename order, input validation, per-file OCR snapshots, duplicate stems, unique output folders, invariant CSV figures, EP scope rows, blank incomplete metrics, formula-safe quoting and atomic updates that preserve an earlier report on failure.

`tests/BatchUiStateTests.cs` passed 35 checks without starting Word. Changing the folder clears the old input selection and disables Start batch until the new folder is loaded; completed results stay accessible. Tests also covered quoted folder paths, moving/removing rows, PDF-only OCR choices, and recoverable CSV saving when the report is locked.

`tests/BatchIntegration.ps1` passed 117 actual hidden WinForms/Word checks using synthetic documents. Four files produced all six comparisons in the requested direction and sequence, with independently expected word counts, crossover, yellow insertions and report rows. EP returned three scope rows with expected insertion counts and excluded translated claims. Uncertain language paused before producing output and an explicit current-pair retry returned 80.0% / 5 new words / 1 insertion. Known language mismatches and a corrupt input failed individually while valid pairs continued. Cancelling an active or paused batch and skipping an unconfirmed pair retained completed outputs and saved explicit incomplete statuses. The two tabs prevented concurrent jobs, source hashes stayed unchanged and temporary jobs were removed. The interface was rendered at its normal and 760 x 580 sizes.

The existing Single comparison tab also passed its actual UI regression: initial short-text detection stopped without output, then explicit manual confirmation produced 80.0% crossover, 5 new words and 1 insertion. Both attempts reset their controls, removed temporary jobs and preserved source hashes.

The pilot plan has 34 scenarios and the fresh template has 49 Not run rows. The team's real-document/device cases remain for the pilot group to run; these developer checks do not mark them passed. Controlled fixture contents and reference figures are unchanged, and short examples still require manual language checking when detection is uncertain. Installation preserves an existing editable pilot log, so newly added batch cases are available in the fresh release template.

## Version 1.0.11 — 5 October 2026

Added an automatic source-language check using the Windows local ELS detector on final main-document text. Word proofing metadata and the OCR recognizer's selected language are not treated as detected source language. Readable DOCX pairs are checked before Word starts where their scopes can be resolved; all formats are checked again after finalization/conversion and before Word comparison. EP Description and selected original-language Claims are checked separately, excluding translated claims and other cropped material.

Detected language mismatches stop without publishing a comparison or figures. Short, mixed-language, unclear or unavailable detection also stops for manual review. The **Language checked — compare** action confirms matching relevant source languages for one retry; it cannot override a detected mismatch. Manual confirmation is noted beside the same three requested statistics. Empty WF inputs retain their existing behaviour. Detection uses agreement across bounded, distributed text samples and can abstain; it is not a guarantee that every language or language change will be identified correctly.

`tests/LanguageDetectionTests.cs` passed 28 native detector checks across English, German, French, Spanish, Italian, Dutch, Russian, Japanese and Chinese, including Unicode normalization, short/technical text and mixed/distributed/minority passages. `tests/EngineTests.cs` passed 31 final-text, revision, highlight and external-link checks. `tests/LanguageGuardIntegration.ps1` passed 13 worker cases: early DOCX mismatch/manual-review stops, misleading proofing metadata, manual retries, EP translated-claim exclusion and original-claim mismatch, and post-conversion PDF mismatch. A confirmed mismatch remained blocked. Failed jobs published no outputs or saved comparison, source hashes stayed unchanged and each owned Word session closed.

All 11 controlled WF/EP reference comparisons still passed with the explicit same-language confirmation documented for their short synthetic text. Counts, crossover, source hashes and private Word cleanup remained correct. The manual-review interface was rendered and checked at the default and 760 x 580 sizes; its action and instructions remain accessible with scrolling. The pilot plan now has 27 scenarios and the blank template has 40 Not run rows, including language mismatch and uncertainty checks. Existing completed logs remain preserved on installation.

The actual hidden desktop UI flow also passed: ordinary Compare submitted no manual confirmation and stopped without output; the review action submitted confirmation and returned 80.0% crossover, 5 new words and 1 insertion with the review note. Both attempts reset their controls and removed their temporary job folders; source hashes stayed unchanged.

## Version 1.0.10 — 5 October 2026

Added a Desktop button beside the results folder and allowed comparison/summary output to the actual Windows Desktop and its descendant folders, including a Desktop managed by OneDrive. The default results folder is unchanged. UI and worker use the same output validator; private temporary working folders keep their existing local restrictions. Other known OneDrive output folders and network locations remain unsupported. Results saved to a OneDrive-managed Desktop may sync according to that user's OneDrive settings; comparison and OCR processing remain on this PC.

`tests/DesktopOutputPathsTests.cs` passed 14 checks for Desktop paths, case/normalization, sibling-prefix and parent-directory escapes, other OneDrive output folders, separate temporary-folder rules and UNC rejection. No mapped network drives were available for its conditional mapped-drive check. `tests/DesktopOutputIntegration.ps1` saved a synthetic comparison and its summary directly on the actual OneDrive-managed Desktop in desktop Word: 5 new words, 1 inserted word, 80.0% crossover and yellow insertion highlighting. Both source hashes stayed unchanged, the private Word session closed and the two synthetic Desktop result files were removed. The updated interface was rendered and inspected; the Desktop/Browse controls and destination field were readable.

README, acceptance checklist, pilot plan in Word/Markdown and P20's expected result in the blank pilot log now cover Desktop output and OneDrive syncing. Installation continues to preserve an existing completed pilot log.

## Version 1.0.9 — 5 October 2026

Corrected the final-text reader's handling of tracked paragraph-mark deletions. A deleted paragraph mark joins the surviving text to the following paragraph; it must not create an artificial word or fingerprint boundary. This could previously withhold crossover despite matching new/comparison text. The existing text and word-count reconciliation checks remain active, and the saved comparison retains its tracked revisions.

`tests/EngineTests.cs` passed 28 checks, including consecutive paragraph joins, joined word fragments with insertion flags, ordinary paragraph/cell boundaries, genuine changed or reordered words, historical property snapshots and revision preservation after highlighting. The portable EP suite passed 50 checks. Actual desktop Word tests passed paragraph joining and splitting (4 words, 0 insertions, 100% crossover) and the reported EP job now reconciles in all three scopes. Its new/insertion counts stayed unchanged, yellow insertion highlights and deletions remained present, both source hashes stayed unchanged and each private Word session closed.

All 11 existing WF/EP pilot reference comparisons also passed against this version, including native moves, formatting/run boundaries and the separate/combined EP edits. Expected figures, unchanged sources and private Word cleanup remained correct.

## Pilot materials — 2 October 2026

Added a practical 25-scenario pilot plan in Word and Markdown, an editable CSV log, and 15 synthetic Word fixtures. The log includes separate EP scope rows and an optional combined-edit cross-check: 38 rows, all starting as Not run with blank actual metrics. The plan covers real EP/WF Word/PDF families, scopes, OCR, OneDrive, local installation, cancellation, source preservation, IT review and criteria for pilot entry/wider rollout. These materials are included in the team ZIP and a standalone `PilotTestMaterials.zip`.

All 11 supplied reference comparisons passed on this PC in actual desktop Word: six WF cases (identical, insertion, deletion, replacement, formatting/run boundaries and moves) and five EP cases (identical prepared, uncropped/prepared, description-only edits, claims-only edits and both sections edited). Expected scope counts/crossover matched, source hashes remained unchanged and each verified private Word session closed. This verifies the example reference values; it does not mark the team's pilot runs as completed. Representative real team documents and device configurations remain to be tested by the group.

Pilot instructions/fixtures are source-built with `tests/Build-PilotMaterials.py`. The application has no new Python runtime dependency. Installation preserves an existing editable pilot log, and uninstall leaves that log; users should save their completed log copies in an approved separate location.

## Version 1.0.8 — 2 October 2026

Replaced the desktop/window/executable icon with a custom document-comparison symbol: two paper sheets and opposing arrows in Galaxy Purple, Hillier Pink and white. The icon was drawn as local vector geometry and exported with transparent 16, 24, 32, 48, 64, 128 and 256 pixel frames. The smallest frames use fewer details and thicker strokes. Source SVG, PNG, ICO and rendering script are in `src/assets`; the application header still uses the original RWS artwork.

The actual icon frames were inspected on light and dark backgrounds and independently reviewed for recognisable documents and comparison arrows. ICO dimensions and transparency were verified. The desktop shortcut uses a separate versioned ICO filename so Windows can select the new artwork without clearing the icon cache or restarting Explorer. The installer includes that exact file in its copying and managed uninstall list.

## Version 1.0.7 — 2 October 2026

Updated the desktop appearance using the supplied May 2025 RWS office-signage guide: original primary full-colour logo with black wordmark on white, Galaxy Purple `#3E016F`, Hillier Pink `#E60054`, bold headings, neutral borders and purple EP table headers. The executable, window and desktop shortcut use the original full-colour logomark. Artwork is embedded in the executable; the source guide, reference screenshots and OCR text are not distributed. No font or online asset dependency was added.

Articulat CF is inferred from the guide's embedded font names and used where installed. It is not installed on this test PC, where Arial is used as a fallback. Fonts were not extracted or redistributed. The guide covers physical signage, so this is an adaptation of its visual principles for the app rather than a separate digital-interface specification. Artwork and colour provenance are recorded in `src/assets/README.md`.

The .NET Framework build passed. The actual Windows desktop interface was rendered and reviewed in WF ready/results states, EP with all three scope rows, and the 760 x 580 minimum window size with scrolling. Logo proportions, selector text, table headers, figures and result actions were readable. A clipped lower button border found during review was corrected and the changed screens were verified again. Purple/white contrast is approximately 14.6:1; grey supporting text on white is 7.5:1. The native selector displays correctly; `DrawToBitmap` can omit its text in a preview capture. Additional team DPI/font configurations remain part of the rollout checklist.

## Version 1.0.6 — 2 October 2026

Added the WF/EP selector. WF retains its whole-document comparison and three figures. EP compares Description plus the first original-language Claims section and reports those figures for Description, Full specification and Claims. All section insertion counts describe the yellow insertions in the same saved comparison. Covers, separate drawings, translated claims, references and running headers/footers are excluded from EP; section headings do not add to its counts.

`tests/EpSectionsTests.cs` passed 61 checks covering section extraction, languages, original/translated claim selection, notes, revisions, moved text, source preservation, and refusal of missing, repeated, empty or nested boundaries. The supplied uncropped and prepared B1 documents have identical text fingerprints for all three scopes. Their actual Word comparison returned 100% crossover and no insertions: Description 6,074 words, Full specification 6,350 words, Claims 276 words. Both source hashes remained unchanged; derived sample test files were removed locally.

`tests/EpIntegration.ps1` verified description additions/replacements (10 words, 2 inserted, 80.0%), claims edits (17 words, 6 inserted, 64.7%), their full specification (27 words, 8 inserted, 70.4%), preserved yellow insertions and tracked deletions, and refusal of repeated claims headings. A moved paragraph retained native move-from/move-to markup and returned 100% crossover in all three scopes (Description 90 words, Full specification 96, Claims 6), with no insertion words. `tests/SimpleResults.ps1` confirmed WF still saves exactly its original three summary lines.

`tests/EpPdfIntegration.ps1` passed actual Word conversion/comparison of synthetic native PDFs with images confined to excluded drawing pages, without OCR. Description returned 9 words/1 insertion, Claims 7/1 and Full specification 16/2. Images retained within Description correctly required OCR and produced no incomplete output. Successful and refused jobs preserved input hashes and closed their verified private Word processes. UI compilation and a rendered three-scope result table were checked locally.

EP requires recognisable standalone section headings and a numbered first claim. Word PDF conversion can merge a heading with nearby text; such files require preparation. Whole-PDF OCR still rejects pages returning no text, including excluded picture-only drawing pages; use a prepared PDF or verified Word transcription. The supplied DOCX pair was tested; no claim is made that every EP publication layout has been validated.

## Version 1.0.5 — 2 October 2026

Final results now contain only crossover percentage, total new-file word count, and highlighted insertion word count alongside the comparison document. Discount and page estimates have been removed from the app, worker result and saved summary. Historical tests for those retired features are under `tests/retired`.

`tests/SimpleResults.ps1` passed an actual Word comparison: 80.0% crossover, 10 total words and 2 highlighted insertion words. The ordinary saved summary contains exactly those three lines. Existing checks also confirmed yellow insertion highlights, unchanged source files, private Word cleanup, and absence of discount/page fields from the worker result. Conversion or reconciliation problems can still add relevant review notes.

## Version 1.0.4 — 29 September 2026

Discount now has a 0% minimum in both the app and saved summary. This replaces the negative-discount behaviour in 1.0.3. Whole-percent rounding remains.

All 12 cases in `tests/DiscountTests.cs` passed, including negative calculations capped at zero, the 45% maximum with no insertions, positive rounding, empty/unreconciled content, and worker/UI JSON transfer.

## Version 1.0.3 — 29 September 2026

Added discount = 45 - (100 × highlighted insertion words / total words in the new document), displayed as a whole percentage in the app and saved summary. Negative values are retained. Empty or unreconciled new documents show discount unavailable.

`tests/DiscountTests.cs` passed 10 cases covering the formula, positive/negative rounding (halfway away from zero), zero insertions, all insertions, zero total words, unreconciled content, and worker/UI JSON transfer. `tests/DiscountIntegration.ps1` passed an actual Word comparison: 2 inserted words out of 10 total produced 25% discount in both the worker result and saved summary. Existing crossover, insertion counts, source preservation and Word cleanup checks also passed.

## Version 1.0.2 — 29 September 2026

OneDrive folders can now supply archive and new input files. Comparison and OCR still use local working copies; the existing local-output restrictions remain.

`tests/OneDriveInputs.ps1` passed with a simulated OneDrive root: both UI and worker input validation accept the documents, local copying preserves source contents, and OneDrive output/working folders and UNC inputs remain rejected. No files were uploaded. Live Files On-Demand downloading was not exercised; availability depends on the OneDrive client and connection.

## Version 1.0.1 — 29 September 2026

The page estimate now measures highlighted insertion words / 500 and displays one decimal place in the app and saved summary. The earlier whole-page allowance has been removed.

`tests/InsertionPages.ps1` passed actual Word comparisons: 1,125 total words with 625 highlighted insertion words produced 1.3 pages of new content; an unchanged 500-word document produced 0.0 pages. Both checks also verified word counts, crossover, source preservation, and private Word process cleanup.

## Initial build — 28 September 2026

Synthetic files only; no confidential source documents were used.

Passed on this computer with desktop Word 16 and Windows local OCR (en-GB):

- 13 focused checks for word boundaries across formatting, partial-word edits, inserted/deleted table rows, move markup, headers/notes, yellow highlighting, preserved revisions, cached fields, and detached external links.
- Actual Word comparisons: additions/replacements (10 words, 2 inserted, 80% crossover), identical documents (100%), deletion-only edits (100%), unrelated wording (0%), and an empty new document (undefined crossover).
- Output reopened in Word with tracked changes intact. Source-file hashes remained unchanged.
- A 90-word document with a moved paragraph retained native move-from/move-to markup and reported 100% crossover.
- Existing output was refused and its hash remained unchanged.
- Native text PDF comparison reported the expected 80% crossover.
- A mixed selectable-text/scanned-page PDF passed local OCR and Word comparison; identical recognised inputs reported 100% crossover.
- Image-bearing PDF without OCR was refused. A PDF page with no recognised text prevented publication of a misleading comparison score.
- OCR helper checked page boundaries, blank pages, malformed PDFs, output conflicts, source preservation, and content-free errors.
- UI cancellation stopped its worker and verified private Word process, removed its temporary folder, and restored the controls.
- Desktop interface rendered and was visually inspected.

The full Word tests require a normal Windows desktop execution context. The restricted coding environment could not start Word correctly, so integration tests ran outside that restriction. No network service was used by the comparison/OCR implementation.

This is a pilot build, not an organisation-wide certification. Test the team's representative files, Word builds, device policies, high-DPI displays, signed-package requirements, and confidentiality controls using `tests/AcceptanceChecklist.md` before broad deployment. OCR spelling and reading order require human review. The package is unsigned.

## Reproduce from source

Run `build.ps1` in Windows PowerShell under your organisation's approved script policy. The build uses the installed .NET Framework C# compiler and downloads no dependencies.

Run these test scripts from the source directory in the normal desktop context:

- `tests/Integration.ps1`
- `tests/PdfIntegration.ps1`
- `tests/MovesAndOverwrite.ps1`
- `tests/UiCancellation.ps1` (Windows PowerShell in STA mode)
- `tests/EpIntegration.ps1` (add `-IncludeSamples` only when the supplied local sample pair is available)
- `tests/EpPdfIntegration.ps1`
- `tests/BatchIntegration.ps1` (runs synthetic batch UI/Word cases; use `-SkipBuild` with a current release)

`tests/BatchPlanTests.cs` and `tests/BatchUiStateTests.cs` use the .NET Framework references from `build.ps1`; compile them with the application sources and select the corresponding test class with the compiler's `/main` option. The UI state checks do not start Word.

`tests/EngineTests.cs` compiles with `src/DocxContent.cs` and the .NET Framework XML/ZIP references listed in `build.ps1`. `tests/OcrSmoke.ps1` checks the OCR helper independently. Test artifacts contain synthetic text only.

`tests/EpSectionsTests.cs` compiles with `src/DocxContent.cs`, `src/EpSections.cs` and the same XML/ZIP references. It uses synthetic fixtures and optionally validates the supplied local DOCX pair when present; its temporary documents are removed after the run. EP sample integration also removes its derived user-content test files after successful verification.
