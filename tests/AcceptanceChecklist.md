# Desktop rollout acceptance checklist

Use synthetic documents without confidential information. Record the application version, Windows version, Word build/bitness, installed OCR languages, tester, and date. Check both the result shown in the application and the saved `.docx` in Word. A passing static or automated build check does not replace these interactive Word/OCR checks.

## Appearance

- [ ] Confirm the primary full-colour RWS header logo with black wordmark appears on white, keeps its aspect ratio, and has clear surrounding space. Check the app and installed desktop icon show the distinct two-document comparison symbol with purple/pink opposing arrows, readable on light and dark backgrounds.
- [ ] Confirm Galaxy Purple buttons/headings, Hillier Pink accent, white backgrounds and neutral borders match the supplied signage guide. Check enabled, disabled, hover, focus and selected states remain readable.
- [ ] Where Articulat CF is installed, confirm the app uses it. On other PCs, confirm the Arial fallback renders correctly without font downloads.
- [ ] Check WF and all three EP result rows at normal and high-DPI scaling, the minimum window size, and maximised. Confirm controls and result actions remain accessible by scrolling and keyboard.
- [ ] Confirm Browse beside Archive and New is easy to find, supports the accepted file types, and fills the correct input. Cancelling a dialog leaves the previous selection intact; drag/drop and entered paths still work. Browse the batch input folder and each results folder as well.
- [ ] On each tab, select files, complete a job, then use Reset and start another job without closing the app. Confirm selected inputs, displayed results and manual language confirmation are cleared; job type and save folder remain selected; saved DOCX/TXT/CSV outputs remain unchanged. Check Reset after single-job uncertain language, completed/failed/cancelled work and an already-empty tab. It must remain disabled during active processing and batch language/report pauses; a prior manual confirmation must not carry into a newly selected job.

## Text and statistics

Confirm **WF** is selected by default and retains the existing whole-document behaviour. The result screen and saved summary must report only the three requested statistics: **Crossover**, **Total word count of new file**, and **Highlighted insertion word count**. Both counts must be word counts. Run the following fixtures in WF.

The short synthetic examples may initially return Unknown/Mixed language. Confirm that both relevant texts are English, then use **Language checked — compare** to run the metric check. Record the initial stop and manual confirmation; the reference counts do not change.

| Fixture | Expected result |
| --- | --- |
| Archive and new both contain `Alpha beta gamma delta.` | 4 new words, no inserted words, 100% crossover; no new yellow insertions. |
| Archive contains `Alpha beta gamma delta.`; new contains `Alpha beta gamma delta epsilon.` | 5 new words, 1 inserted word, 80% crossover; `epsilon` highlighted yellow. |
| Archive contains `Alpha beta gamma delta.`; new contains `Alpha beta gamma.` | 3 new words, no inserted words, 100% crossover; deleted `delta` remains a tracked deletion in All Markup. |
| Archive contains `Alpha beta gamma.`; new contains `River stone cloud.` | 3 new words, 3 inserted words, 0% crossover; replacement text highlighted and archive text retained as tracked deletions. |
| New is empty and archive contains text | 0 new words, 0 inserted words; crossover unavailable, with no division-by-zero error; deletions retained. |
| Archive is empty and new contains `Alpha beta gamma.` | 3 new words, 3 inserted words, 0% crossover. |

Use longer natural passages as well as these small fixtures: Word's decisions on short or ambiguous changes can differ from its decisions on paragraphs. If a fixture's revision classification differs, examine the Word revision structure before changing the asserted metric.

- [ ] Move a distinctive multi-sentence paragraph, including its paragraph break. Confirm the output preserves Word's move revisions and its final text matches the new source. Recognised moved-to words contribute to new words but not inserted words; moved-from text does not inflate the new count. Moved paragraph-break revisions must not falsely withhold crossover when strict final-text and word-count checks reconcile. If Word classifies the passage as delete/insert, the percentage follows that classification and the limitation is understood.
- [ ] Add a second copy of an existing paragraph. Confirm the additional occurrence counts as inserted when Word classifies it as an insertion.
- [ ] Replace a single word and edit part of a longer word. Confirm a changed word contributes one inserted word, not a count of inserted fragments.
- [ ] Split a word across bold/italic runs. Confirm it is counted once.
- [ ] Change only bold, font size, or colour. Confirm this does not add inserted words.
- [ ] Join `Alpha beta.` and `Gamma delta.` into one paragraph with a space between the sentences, then test the reverse split. Confirm 4 new words, 0 inserted words and 100.0% crossover, with paragraph changes preserved as tracked revisions.
- [ ] Check case-only edits, punctuation-only edits, apostrophes, hyphens, numbers, accented text, and non-Latin text. Confirm the result follows the documented word-token rules and document any language limitations.
- [ ] Check table text, text boxes, headers, footers, footnotes, and endnotes. Confirm counted insertions receive yellow highlighting and that repeated page headers are not counted once per displayed page.
- [ ] Include comments and fields. Confirm comments and field instructions do not inflate word counts. Validate visible field results manually.
- [ ] Confirm every result with a reported percentage has matching final-content text and word counts between the new document and comparison. If reconciliation fails, confirm the percentage is withheld and a warning is visible.

## Source-language guard

- [ ] Compare sufficiently long, clearly English archive text with clearly German or French new text, then test the reverse order. Confirm a known mismatch stops before Word comparison, reports the language problem and publishes no comparison figures. Repeat with an EP mismatch confined to Description, then to original-language Claims; excluded translated claims must not cause a mismatch.
- [ ] For a known single-job mismatch, confirm a clear pop-up names the affected scope and detected Archive/New languages, explains that the inputs must use the same source language, and states that no comparison was produced. Dismiss it, choose matching-language inputs and compare again. Unknown/Mixed must retain the existing manual-review behaviour instead of this known-mismatch pop-up.
- [ ] For readable DOCX pairs with resolvable scopes, confirm language checking happens before Word starts. For DOC, text PDF and scanned PDF inputs, confirm conversion/OCR can run first but a mismatch still stops before Word comparison. Original inputs remain unchanged.
- [ ] Give English text a German proofing-language setting, or use incorrect proofing metadata in the other direction. Confirm detection follows actual text and never treats the proofing setting as authoritative.
- [ ] Use short or mixed-language text that produces Unknown/Mixed. Confirm the first attempt stops with manual-check guidance and exposes **Language checked — compare**. Check matching relevant source languages before retrying; confirm the manual confirmation appears in review notes without adding a fourth statistic.
- [ ] Confirm a known mismatch cannot be overridden by manual confirmation. Check the EP Description pair and selected original-language Claims pair separately before confirming an uncertain result.
- [ ] Confirm an empty WF input is not rejected merely for having no language; retain the documented empty-document metrics. Confirm unavailable detection or uncertain text requires manual review rather than a guessed language match.

## EP scopes

- [ ] Select EP and compare prepared specifications with clear Description and original-language Claims headings. Confirm one `[new base name] vs [archive base name].docx` and its companion summary, with three sets of crossover, new-text word count, and highlighted insertion word count: Description, Full specification, and Claims.
- [ ] Use known text counts. Confirm Description and Claims include only their respective content and Full specification equals their sum. Confirm section titles do not add to either word count.
- [ ] Add cover text, drawing labels, translated claims, and a reference bibliography to create uncropped publication fixtures. Confirm these contents are omitted from the EP DOCX and do not contribute to any EP scope's word count or insertion count. Confirm the two retained section headings do not add words to the statistics.
- [ ] Change Description only, then Claims only. Confirm each scope's statistics follow its own words and insertions; Full specification reflects both scopes together.
- [ ] Move a distinctive passage between Description and Claims. Inspect the single Word comparison. If Word recognises a move, confirm its moved-to words remain retained in the destination scope; do not treat them as inserted merely because they changed sections.
- [ ] Test missing, duplicate, and ambiguous section headings independently in archive and new inputs, along with empty sections and an unrecognisable first numbered claim. Confirm EP stops with useful preparation guidance and never substitutes whole-document statistics or guessed boundaries.
- [ ] Test standalone `Description` / `Beschreibung` headings with `Claims`, `Ansprüche`, `Patentansprüche`, and `Revendications`, including different letter case and a final colon. Confirm the first recognised claims section following Description is selected, rather than always selecting English; later translated claims are excluded.
- [ ] Put section headings inside tables, text boxes, headers, and footers. Confirm these unsupported boundaries do not produce guessed EP statistics. Prepare the headings as standalone body paragraphs and retry.
- [ ] Within a clearly headed Description section, include a normal data table with first-row labels `Abbreviation`, `Component Type`, `Description` and populated rows below. Confirm the Description column label does not stop EP processing and that the entire table contributes to Description and Full specification counts. Nested original/translated Claims headings and ambiguous section headings must still stop processing.
- [ ] Include non-breaking spaces split across Word text runs, including runs consisting solely of a non-breaking space. Compare an EP file with itself and confirm all three scopes show 100.0% crossover and zero insertions, with words remaining separated in the saved comparison. Check narrow non-breaking spaces too; the original file must remain unchanged.
- [ ] Test first-claim numbering `1.`, `1)`, `(1)`, and Word automatic numbering. Confirm a bare `1` followed by text is refused until prepared with supported numbering.
- [ ] Add headers and footers to EP fixtures. Confirm their text is omitted from the comparison and all three counts. Confirm tables, inline graphics, and text describing drawings within Description remain in their appropriate retained sections.
- [ ] Confirm both prepared and uncropped native Word inputs preserve their originals. Choose EP with a PDF New input and confirm the WF recommendation is visible; check the general guidance for DOCX files converted from PDF. Prepare copies of both inputs to the same intended scope and choose WF, removing covers/translations/separate drawings where appropriate. WF counts all included text and gives one row; native Word EP keeps three scopes. No automatic mode switch should occur.
- [ ] In EP, compare a text PDF with images confined to excluded cover or drawing sections. Confirm those images alone do not require OCR. Add an image within Description or Claims and confirm the conservative OCR requirement still applies there.
- [ ] In EP with OCR, include excluded blank or picture-only pages. Confirm the existing whole-PDF OCR completeness check still stops processing, and prepared-PDF or verified-Word guidance is useful.
- [ ] Switch back to WF and confirm the same files are assessed as whole documents rather than retaining EP scope exclusions.

## Multi Archive

- [ ] Select one new file and four distinct archives. Confirm exactly four comparisons are planned, each with the same New and a different Archive; no archive-versus-archive comparisons run. Check a single archive is also allowed.
- [ ] Add multiple archives with the file picker and drag/drop, including from different accessible folders. Re-add an existing archive and confirm it is not duplicated and its OCR choice remains unchanged. Load a folder and confirm supported top-level files replace the list and the selected New file is excluded. Remove an unwanted archive and confirm the comparison count changes.
- [ ] Choose scanned PDFs independently for New and archives. Confirm each pair receives the correct two OCR choices, DOC/DOCX OCR stays unavailable, and changing New clears its previous OCR choice.
- [ ] Use Full Document archives with independently known crossover figures. Confirm highest verified crossover appears first in both the UI and multi-archive-results.csv, the summary names that archive, and each successful pair has a Word comparison/TXT summary with the expected counts and yellow insertions.
- [ ] Use EP archives where the highest Description crossover belongs to a different archive from the highest Full specification crossover. Confirm ranking follows Full specification, all three scope rows remain visible, and the CSV repeats the same rank beside all three rows for that archive.
- [ ] Include tied scores, zero crossover, unavailable crossover despite available counts, empty new text, failed and unfinished comparisons. Confirm exact score ties share a rank, valid zero is ranked, invalid/unavailable/incomplete results are unranked, and no best-match claim appears if no verified percentage is available. Ranking uses full precision; displayed crossover remains one decimal place.
- [ ] Include a known source-language mismatch before valid archives. Confirm it fails without output, later valid matches continue, and one Multi Archive mismatch summary appears after the search. Check Unknown/Mixed pauses only the current pair; manual confirmation/Skip archive do not carry confirmation to the next archive.
- [ ] Cancel after one completed comparison and confirm that file and its summary remain unchanged, unfinished archives have explicit incomplete statuses, and the best-match summary describes completed comparisons only. Lock/retry the CSV and confirm completed comparisons are preserved without being rerun.
- [ ] Confirm Reset clears the fixed new file, archive selection/OCR and ranked results while retaining job type, destination and saved outputs. Check minimum-width layout and that all three tabs prevent concurrent jobs; closing an active search must retain completed outputs and stop only owned processes.

## Word output integrity

- [ ] Open the result in Word without a repair prompt or corruption warning.
- [ ] Select All Markup; confirm inserted text is yellow and deletion/move revisions remain inspectable.
- [ ] Inspect pre-existing coloured highlights. Check that inserted words become yellow and the expected original formatting is otherwise retained.
- [ ] Compare inputs with existing tracked changes, including unresolved deletions and insertions. Confirm the tool compares their final text by accepting input revisions only in temporary copies. Confirm the originals retain their existing revisions and the output retains the new comparison revisions.
- [ ] Inspect author names, comments, document properties, and retained deleted text. Confirm users understand that comparison output is not sanitised.
- [ ] Confirm the saved filename is `[new base name] vs [archive base name].docx`. Test equal base names, non-ASCII names, very long names, and an existing result with the same proposed name. Confirm the tool does not silently overwrite inputs or previous results.
- [ ] Open the result from both the application and its folder. Confirm any companion summary agrees with the result screen.
- [ ] Compare source-file hashes before and after comparison; confirm neither original changed.

## Batch comparison

- [ ] Paste and browse a folder, then load files. Confirm only top-level DOC, DOCX and PDF files appear, Word `~$` files and subfolder contents are excluded, and numbered names sort naturally. Check Move up, Move down and Remove; the visible order must determine which input is archive and which is new.
- [ ] Run four WF files in a verified order 1, 2, 3, 4. Confirm exactly six pairs in this order: 2 vs 1, 3 vs 1, 4 vs 1, 3 vs 2, 4 vs 2, 4 vs 3. Check each completed comparison and TXT summary against its figures and one completed data row in `batch-results.csv`. Repeat with representative paragraph joins/moves and confirm crossover remains available for every pair whose final text and word counts reconcile; investigate genuine mismatches individually.
- [ ] Repeat in EP. Confirm each completed pair has three report rows, Description/Full specification/Claims, with full word and insertion counts equal to their section sums. The report contains no extra page or discount metric.
- [ ] Use different input formats with equal filename stems and rerun a batch into the same parent. Confirm unique batch folders and numbered pair subfolders preserve every result without collisions or source modification.
- [ ] Include a corrupt file and a known source-language mismatch. Confirm failed pairs have error notes and blank metrics, while other pairs continue. Include Unknown/Mixed text; confirm the current pair pauses, Language checked — continue applies only to the manually checked pair, Skip pair leaves it without success metrics, and a known mismatch remains blocked.
- [ ] Include multiple known language-mismatch pairs in a batch. Confirm other pairs continue without one dialog per failed pair and one summary pop-up identifies the mismatches when the batch finishes. Repeat with Cancel batch after a mismatch; completed outputs and report remain. Closing the app must not open a summary dialog. Unknown/Mixed remains a manual-review pause and is not presented as a detected mismatch.
- [ ] Cancel after at least one pair completes, including during Word/PDF/OCR work. Confirm completed DOCX/TXT files and report results remain, unfinished pairs are not labelled completed, unrelated user Word work remains intact, and the interface can start another job.
- [ ] Use downloaded OneDrive folder inputs and per-PDF OCR settings. Save to the default local parent, Desktop, a Desktop subfolder and another approved OneDrive folder. Confirm all accept Batch folders, comparisons, summaries and CSV; OneDrive results may sync while private working files remain local.
- [ ] Load ordered inputs from a UNC share and a mapped network drive and save the batch to each accessible destination. Check all completed pair metrics and CSV entries. Using synthetic files under IT supervision, test a read-only destination, unavailable share, interrupted connection and locked CSV: truthful statuses, preserved completed work and report retry controls must remain usable.
- [ ] Check both comparison tabs at minimum and high-DPI sizes, with keyboard navigation. During a batch, confirm ordering/options remain consistent and the manual-review controls and result actions are reachable.

## PDFs and OCR

- [ ] Compare selectable-text PDF/PDF and mixed PDF/Word pairs.
- [ ] Compare a scan-only PDF with OCR enabled. Check recognised text against a known transcription, including names and numbers.
- [ ] Compare a mixed PDF containing both text and scanned pages. With OCR enabled, confirm all pages are represented and page order is correct.
- [ ] Run a scan-only PDF without OCR. Confirm the tool requires OCR or valid prepared sections rather than reporting a misleading percentage. In WF, test an ordinary text PDF with a logo and confirm the conservative image check also requires OCR, as documented.
- [ ] Test multi-column pages, tables, line-end hyphenation, rotated pages, poor scans, and a large PDF. Record whether conversion quality is adequate for team use.
- [ ] Test unsupported/missing OCR language support. Confirm a useful error and no misleading success result.
- [ ] Include a genuinely blank PDF page and a page containing only a picture. Confirm a page with no recognised text stops comparison and the guidance explains checked transcription or removing genuinely blank pages from a copy. Confirm incomplete page-processing status also stops comparison.
- [ ] Confirm OCR does not silently truncate large pages or long documents; sample text from the first, middle, and last pages.
- [ ] Confirm the OCR option clearly explains that full-PDF text recognition loses original page layout.

## Desktop operation and failures

- [ ] Test from a standard non-administrator account on a clean team machine.
- [ ] Run the portable folder, install with `Install.cmd`, start from the desktop shortcut, and uninstall with `Uninstall.cmd`. Confirm uninstallation retains saved results and that the installed README's acceptance-checklist link works.
- [ ] Confirm the application remains usable on a 1366x768 display and at the team's 100%, 125%, and 150% scaling settings. Verify that scrolling or resizing reveals all fields, warnings, and result buttons.
- [ ] Confirm startup/first-run handling on a machine where Word has not completed activation or its initial setup.
- [ ] Leave an unrelated unsaved document open in Word before starting, completing, failing, and cancelling a comparison. Confirm that document and its Word session remain intact.
- [ ] Test cancellation while opening Word, converting a PDF, running OCR, comparing, and saving. Confirm the UI remains usable, no partial output is presented as success, and cleanup failures identify the remaining working directory.
- [ ] Test an output folder without write permission, a read-only input, a locked input, a missing input, a corrupt file, and insufficient disk space in a controlled environment. Confirm actionable errors and preservation of originals/results.
- [ ] Test password-protected and organisation-restricted documents. Confirm the tool does not bypass controls and does not leave an unresponsive background Word process.
- [ ] Test a selected output path matching an input path and a duplicate result path. Confirm safe handling before any write.
- [ ] Test two launches or overlapping comparisons. Confirm job folders remain separate and cancelling one job cannot close unrelated Word sessions.
- [ ] Test source/output paths containing spaces, apostrophes, non-ASCII characters, and long paths. Confirm sensible failure messages where platform limits apply.

## Confidentiality and IT review

- [ ] Confirm the release contents and review application code for network calls, telemetry, content logging, and shell command construction.
- [ ] Validate operation with network access blocked under IT supervision. Separately assess Office add-ins, Microsoft 365 connected experiences, antivirus/DLP, Windows services, sync, and backup behaviour; do not infer that these are offline merely because the application has no network code.
- [ ] Verify the default output location is `%LOCALAPPDATA%\LocalDocumentCompare\Results`. Confirm the **Desktop** button selects the actual Windows Desktop and that both it and a folder within it accept results, including when OneDrive manages the Desktop. Confirm the comparison DOCX and companion TXT are saved there and users understand that these result files may sync through OneDrive while comparison/OCR processing remains local.
- [ ] Confirm Office macros do not run when opening a source copy and that only temporary copies are passed to Word. Confirm the original input paths are never edited or saved. Review link/field-update behaviour and Protected View handling under the team's policies.
- [ ] Include DOCX and legacy DOC fixtures with external templates, linked images, hyperlinks, and cached fields, using IT-approved test endpoints. Confirm there are no unexpected fetches or field refreshes under the required Office/network policies. Inspect the saved comparison for retained external references before it is opened outside the worker session.
- [ ] Confirm downloaded archive/new files in a company OneDrive folder are accepted and their originals remain unchanged. Check an online-only input with OneDrive connected, and confirm the guidance is useful when a download cannot finish. No comparison/OCR service is used by the app.
- [ ] Confirm accessible UNC and mapped-drive inputs/results are accepted, including single comparisons and batch folders. Verify read/write permissions, originals unchanged, readable DOCX/TXT/CSV outputs and clear failures for inaccessible locations. Approved OneDrive output folders outside Desktop are also accepted and follow sync rules. Private working folders must still be local and outside known OneDrive paths; inspect the actual working location. Shared-drive transfers and destination access/backup rules are separate from local comparison/OCR. Check folder redirection, junctions and other sync products separately.
- [ ] Review the process-scoped PowerShell `-ExecutionPolicy Bypass` argument for setup and OCR. Confirm enterprise policy still applies, no permanent execution policy is changed, and any required signing or allowlisting is in place.
- [ ] Inspect `%TEMP%\LocalDocumentCompare` after success, failure, cancellation, and a controlled forced termination. Confirm normal cleanup and document residual-file handling after crashes. Do not delete a different active job's files.
- [ ] Confirm results and temporary files have expected account-level access controls and are covered by required disk encryption and retention policies.
- [ ] Confirm the UI, README, and rollout guidance do not promise Office-level offline behaviour, secure erasure, or sanitised output.
- [ ] Decide on code signing, software allowlisting, package integrity checks, update distribution, and support ownership before broad deployment.

Record actual pass/fail results and unresolved limitations. This checklist describes required validation; its presence does not mean those checks have been run.
