# Local Document Compare

Version 1.0.21.0 — team package and repository updated on 15 October 2026 after local review of OCR layout preservation. See [validation results](VALIDATION.md) for completed checks and [the acceptance checklist](tests/AcceptanceChecklist.md) for checks on the team's own configurations.

A Windows desktop tool that compares an archive document with a new document using the installed Microsoft Word application. It saves a Word comparison with inserted text highlighted yellow and deletions and moves retained as tracked changes. Alongside the comparison, it reports crossover percentage, total new-text word count, and highlighted insertion word count. **Full Document** (WF) compares whole documents; **EP** reports these figures separately for Description, Full specification, and Claims. Use a single comparison for two files, **Batch comparison** for every ordered pair from a folder, or **Multi Archive** to find the best archive for one new document.

The interface uses the primary RWS logo, Galaxy Purple, Hillier Pink, white surfaces and bold headings, adapted from the supplied May 2025 office-signage guide. The application and desktop icon show two documents with opposing comparison arrows in the RWS colours. Articulat CF is used when installed; other PCs use Arial. Fonts are not downloaded or bundled. Branding artwork is included with the app, so it needs no network connection or access to the source brand PDF.

## Build from source

On Windows with .NET Framework 4.8, run `build.ps1` from PowerShell to compile the app into `dist\LocalDocumentCompare`. Run `package.ps1` to build the complete team ZIP and checksum in `dist`. Building and packaging do not require Python; the repository includes the current generated Word guides and reviewed synthetic fixtures.

The C# source, scripts, branding assets, documentation and synthetic fixtures belong in Git. Build output, local test artifacts, Python caches, private Word/PDF inputs and generated comparison reports are excluded by `.gitignore`. `tests/PilotResults.csv` is an uncompleted release template; keep filled-in pilot logs outside the repository.

The optional real-document EP checks take explicit caller-supplied paths; they do not search a user's desktop. The `EpSectionsTests` executable accepts archive and prepared paths as two arguments. `tests\EpIntegration.ps1` accepts `-IncludeSamples -SampleArchivePath ... -SamplePreparedPath ...`. Use approved local files for those optional checks and keep their content out of Git.

## Requirements

- Windows 10 or 11 and .NET Framework 4.8.
- An installed, activated desktop version of Microsoft Word. Word on the web is not sufficient.
- Permission to run the application and automate Word under your organisation's policies.
- For scanned PDFs, Windows PowerShell 5.1, Windows local OCR support, and an appropriate installed recognition language.

Finish any first-run Word setup and activation before using the tool. Pilot the application on the team's actual Word builds and document types before wider rollout; see [the acceptance checklist](tests/AcceptanceChecklist.md).

For rollout checks, use the [practical test plan](tests/PilotTestPlan.docx), [editable results log](tests/PilotResults.csv), and [synthetic Word fixtures](tests/PilotFixtures). The plan lists 44 scenarios, including Multi Archive, the EP preparation fixes and Reset, with reference figures, representative document types and release criteria. The log starts as Not run; save completed copies outside the installation folder so updates or uninstalling cannot replace them. Installation preserves an existing pilot log; use a fresh copy of the release template for newly added cases. A [plain-text plan](tests/PilotTestPlan.md) is also included.

## Install or run

The branded `DocumentCompare-QuickGuide.docx` is a three-page handout covering single, batch and Multi Archive comparisons, results, language checks and OCR layout reconstruction. The current guide is included in the team ZIP and installed with the app.

Distribute the entire `LocalDocumentCompare` release folder, keeping its files together.

- **Portable use:** extract the folder to an approved local location and run `DocumentCompare.exe`. Create a desktop shortcut to that executable if needed.
- **Per-user installation:** run `Install.cmd` from the release folder. It copies the application to the user's local application data folder and creates a desktop shortcut. Administrator access is not normally needed.

For a shared-drive rollout, put `DocumentCompare-Team.zip` in the team's shared resources folder. Each colleague should download or copy it to their own PC, extract the entire ZIP, open the `LocalDocumentCompare` folder and run `Install.cmd`. Pilot users should close Document Compare and run the new `Install.cmd` to update their existing installation; saved results and an existing pilot log are retained. Updates are manual. Opening `DocumentCompare.exe` runs the app directly; it does not start an installation wizard.

Use `Uninstall.cmd` to remove the installed application files and shortcut. Comparison results are retained. IT should review and approve the package under the organisation's normal software deployment rules; an internally built executable may need signing or allowlisting.

The setup commands and OCR helper invoke Windows PowerShell with `-ExecutionPolicy Bypass` for that process only. They do not make a permanent execution-policy change, and enterprise Group Policy still takes precedence. IT may require signed scripts or an approved deployment method before allowing these components to run.

## Compare two documents

1. Choose **Full Document** for a whole-document comparison, or **EP** for a prepared European patent specification with reliable section headings. Full Document is selected by default. **Use Full Document when an EP new file is a PDF or a Word file converted from PDF:** conversion does not consistently preserve the section boundaries needed for EP's three scopes. Prepare both files to contain the same intended scope, removing covers, translated claims and separate drawings from checked copies where appropriate. Full Document counts all included text and gives one set of figures.
2. Use **Browse** beside **Archive**, drag a file onto its input, or enter its full File Explorer path. Archive is the earlier version.
3. Add **New** in the same way; it is the version to assess. Supported inputs are Word `.docx`/`.doc` documents and PDF files. Local, OneDrive and accessible shared-drive inputs are supported; use a file path rather than a web sharing link.
4. For a PDF containing scans, select that input's local OCR option. This recognises the whole PDF and reconstructs page sizes, approximate text sizes, indents and spacing where possible. Leave OCR off for a normal PDF whose text can be selected.
5. Choose an output folder. The default is `%LOCALAPPDATA%\LocalDocumentCompare\Results`. Click **Desktop**, enter a folder path, or use **Browse**. You can save to local folders, company OneDrive folders or accessible shared drives, including UNC paths such as `\\server\share\folder` and mapped drives such as `S:\folder`. Shared locations require connectivity and write permission; OneDrive results may sync.
6. Start the comparison. The app checks that archive and new text use the same source language before Word compares them. If it reports **Unknown** or **Mixed**, check both inputs yourself and use **Language checked — compare** only after confirming matching relevant source languages. A detected language mismatch remains blocked. Allow any Word conversion or OCR to finish; large files and PDFs can take longer than ordinary Word documents.
7. Open the resulting `.docx` in Word. The filename uses the input names without their extensions: **`[new file name] vs [archive file name].docx`**. A `.txt` summary with the same name is saved alongside it. The result screen and saved summary show the three statistics for the whole document in Full Document, or for each of the three EP scopes.

The archive and new inputs are not saved over. Keep both originals so the result can be checked against them. In Word, select **Review > Display for Review > All Markup** to see deletions and moves as well as yellow insertions. How revision marks appear also depends on Word's review display settings.

If an input already contains tracked changes, the tool accepts them on its temporary copy before comparison. It compares each file's final text, including its existing insertions and excluding its existing deletions. The originals remain untouched, and the newly generated comparison keeps its own tracked changes. Existing highlights are retained except where inserted text is highlighted yellow.

The tool refuses to replace an existing comparison or summary. Choose another output folder, or move the earlier result first. A comparison is stopped after 20 minutes; try a smaller file and inspect the inputs in Word if this limit is reached.

Use **Reset** on either comparison tab to start another job without closing the app. It clears selected inputs, on-screen results and any manual language confirmation, while keeping the selected job type and save folder. Saved comparison documents, summaries and batch reports remain in their folders. Reset is disabled while processing is active or a batch is paused for language review or report saving. After a single job stops for uncertain language, Reset can discard that selection and confirmation instead of retrying it.

Check the result before sharing it. Tracked deletions remain inside the comparison file and can reveal material removed from the new version. Comments, metadata, and pre-existing formatting may also remain; this is a comparison document, not a redacted or sanitised copy.

## Compare a folder in a batch

1. Open **Batch comparison**. Paste a File Explorer folder path in **Folder of documents**, or use **Browse**, then click **Load files**. Local, OneDrive and accessible shared-drive folders are supported. Only `.doc`, `.docx` and `.pdf` files directly inside the chosen folder are loaded; subfolders and Word temporary files beginning `~$` are ignored.
2. Check the visible file order. Use filename prefixes such as **01**, **02**, **03** for versions from earliest to latest. Names are initially sorted in natural order, so a numbered name ending in 2 comes before one ending in 10. Use **Move up**, **Move down** and **Remove** to adjust the list. Filename sorting does not establish the actual document chronology; your visible order determines it.
3. Choose **Full Document** or **EP** for the entire batch. Use Full Document for EP files converted from PDF; prepare all inputs with the same intended scope first. Select **PDF OCR** for each PDF containing scans. Word inputs use the existing editable-text path; the PDF checkbox does not recover text from images inside Word files.
4. Choose the parent folder under **Save batch to** by entering its path, using **Browse** or clicking **Desktop**. Local, OneDrive and shared-drive results folders are supported. Ensure the drive remains connected and you have write permission. Click **Start batch** and keep the selected inputs unchanged while it runs.
5. Review each completed comparison and the batch report. A new, uniquely named timestamped `Batch` folder contains numbered pair subfolders. Each successful pair saves `[new file name] vs [archive file name].docx` and its TXT summary in its own subfolder, so repeated base names cannot replace another pair's output. The batch root contains **`batch-results.csv`**.

Every later file is compared as **new** against every earlier **archive**. For four files in order 1, 2, 3, 4, the six comparisons run as **2 vs 1, 3 vs 1, 4 vs 1, 3 vs 2, 4 vs 2, 4 vs 3**. The number of pairs is `n × (n − 1) / 2`; a larger folder can therefore take substantially longer than a single comparison.

The report has one data row per completed Full Document pair, or three per completed EP pair: Description, Full specification and Claims. The metric columns are **Crossover (%)**, **New file word count**, and **Highlighted insertion word count**. It also identifies the archive/new files, scope, status, comparison path and notes. Failed or incomplete pairs have blank metrics; read their status and notes rather than treating blank values as zero.

A known language mismatch marks that pair **Failed** and the batch continues with other pairs. There is no pop-up for each failed pair; when the batch finishes or **Cancel batch** is used, one summary pop-up identifies any pairs skipped for a known language mismatch. Review their report notes and choose same-language inputs before rerunning them. Unknown/Mixed language pauses the current pair for review. Check that pair's relevant source languages, then use **Language checked — continue** to retry it, or **Skip pair** to continue without its comparison. Confirmation applies to that pair, not all remaining files, and cannot override a known mismatch. EP Description and original-language Claims still require separate checks.

**Cancel batch** stops further work. Already completed comparison documents, TXT summaries and batch-report results remain in their batch folder after cancellation or a later pair error. Inspect the report to distinguish completed work from failures and unfinished pairs. The existing local processing, source preservation, OCR completeness and language safeguards apply to every pair; originals are not saved over.

If the CSV cannot be updated, close it in Excel and use **Retry saving results** to resume. After a stopped or completed batch, **Save results table** can retry saving the table without repeating Word comparisons.

Once the batch has stopped, **Reset** clears its selected inputs and displayed results so you can load another folder. The job type and save folder remain selected, and saved outputs remain available in File Explorer. Save or retry any results table you need before resetting the on-screen list.

## Multi Archive

Use this tab when you have one new document and several possible archive jobs. It compares that same new document against each archive, producing `n` comparisons for `n` archives; archives are not compared with each other.

1. Open **Multi Archive** and choose **Full Document**, or **EP** when separate section figures are needed. For EP PDFs or PDF-converted Word files, follow the Full Document guidance above.
2. Use **Add archives** to select several files from any accessible location, or drop files into the archive list. Duplicate additions are ignored without changing existing PDF OCR choices. Alternatively, browse/paste an archive folder and choose **Load files**: this replaces the archive list with supported top-level files, excluding the new file if already selected.
3. Browse, drop or enter the **New document** below the archive list. Select local OCR if that PDF contains scanned pages. Use **Remove** for unwanted archive entries, including the new file if it was loaded from the same folder. The new document must be different from every archive.
4. Select PDF OCR for each scanned archive PDF and choose the results destination. Local, OneDrive and shared-drive paths use the established source/output rules.
5. Choose **Find best match**. Each successful pair produces `[new name] vs [archive name].docx` plus a TXT summary in a numbered subfolder inside a new **Multi Archive** results folder.
6. Review the ranked table and best-match summary, and open the comparison before choosing an archive. **Full Document** ranks by whole-document crossover. **EP** ranks by **Full specification crossover**, displaying Description, Full specification and Claims figures for each archive.

Ranked matches appear from highest verified crossover to lowest. Exact score ties share a rank; ranking uses the unrounded percentage while display uses one decimal place. **Rank so far** is provisional until the search finishes. Unavailable crossover, empty new text, failed comparisons and unfinished work have no rank, even when word counts are present. If no archive has a verified percentage, the app explains that it cannot identify a ranked best match. The combined **multi-archive-results.csv** follows the same order and repeats each EP archive's rank beside all three scope rows.

Known language mismatches fail individually while other archives continue, followed by one summary pop-up. Unknown/Mixed pauses the current pair for checking; **Language checked — continue** confirms only that pair, and **Skip archive** continues without it. A known mismatch cannot be overridden. **Cancel search** keeps completed outputs and identifies the best match among completed comparisons only. **Reset** clears the new file, archive list, OCR choices and on-screen results while preserving job type, destination and saved files. Single, Batch and Multi Archive cannot run concurrently. Report-save failures retain completed outputs and use the same retry actions as Batch comparison.

## Source-language check

The app detects language from the actual text using Windows' local ELS language detector. It does not use Word's proofing-language setting as evidence of the source language and does not send text to a language service. In Full Document it checks the final main-document text; separate headers, footers and notes do not determine the source language. In EP it checks Description and the selected original-language Claims separately, excluding covers, drawings and translated claims.

A known difference between the archive and new source language stops the job before Word comparison and produces no comparison document or figures. For a single comparison, a pop-up identifies the affected scope and the detected archive/new languages, and asks you to select inputs in the same source language. Check the sources, replace the incorrect selection and rerun. When both inputs are readable DOCX files and their relevant sections can be resolved, this check happens before Word starts. Legacy DOC files, PDFs and OCR inputs first need local conversion or recognition to obtain their text; the language check then happens before Word compares it.

**Unknown** or **Mixed** also stops the initial attempt. Short examples, technical text or multiple languages may need manual review. Read both relevant sources, confirm they use matching source languages, and then click **Language checked — compare** to retry a single comparison, or **Language checked — continue** for the current paused batch pair. For EP, check the Description pair and the original-language Claims pair separately. This manual confirmation cannot override a known mismatch and is recorded in the result's review notes. The result still has only crossover, new-text word count and highlighted insertion word count.

An empty Full Document input has no source language to check; the existing empty-document comparison behaviour is retained. Automatic language detection is a check, not proof that a document is correct: reviewers should still verify the selected language and PDF/OCR text.

## Understanding the numbers

**Crossover** is the share of new words that the comparison treats as retained from the archive. In Full Document, the counts cover the whole document; in EP, each percentage uses the words in that scope:

```text
Crossover (%) = 100 × (new words − inserted words) / new words
```

For example, a new document with 1,000 words and 250 highlighted insertion words reports 75% crossover, 1,000 total new-file words, and 250 highlighted insertion words. A deletion alone does not lower crossover because deleted words are no longer in the new document. Replacing words contributes inserted words. Repeated text added as a new occurrence can also contribute inserted words.

**Total new-text word count** counts the final text in the selected scope, accepting any existing tracked insertions and excluding existing tracked deletions. In Full Document, this is the total word count of the new file. **Highlighted insertion word count** counts the words marked as insertions by this comparison within the same scope. Retained text, deletions, and Word-recognised moves do not contribute to that insertion count. Existing yellow highlighting in a source document does not make retained text an insertion.

These are textual comparison measures, not an assessment of meaning or originality. Word determines the matching and revisions. A moved passage counts as retained when Word recognises it as a move; a passage treated by Word as deletion and insertion contributes inserted words. Case edits can count as insertions when Word marks them as changed words. Formatting-only changes do not add words. When checking a comparison's final text, the app accounts for joined paragraphs and Word's moved paragraph-break revisions; the same strict text and word-count checks still apply before crossover is reported.

Words are counted using letters and numbers, allowing internal apostrophes and hyphens. A word split across formatting runs is counted once. If any character within a word is inserted, that word counts as an inserted word. Punctuation on its own does not count as a word. This count can differ from Word's built-in status-bar count.

In Full Document, the count includes body text, tables, text boxes, headers, footers, footnotes, and endnotes. A stored header or footer is counted once rather than once on each displayed page. Comments, deleted text, and field instructions are excluded. Text embedded in images is not counted unless recovered by OCR. Neither images nor page layout add words.

In Full Document, an empty new document has unavailable crossover rather than 0%. EP requires non-empty Description and Claims sections. If the comparison's final text or word count does not reconcile with the new document, the tool withholds the percentage and reports a warning. Review such a result manually rather than interpreting a missing percentage as zero.

## EP specifications

Use EP for editable Word specifications with clear, independently checked section boundaries. For an EP new file received as PDF or converted from PDF into Word, **Full Document is the recommended workflow** because converted section headings are not consistently reliable. Prepare both inputs to the same intended scope first. Full Document counts the entire included document and does not produce the three separate EP scope rows. A file's DOCX extension does not establish that it originated in Word; check PDF-converted sources yourself.

EP produces one Word comparison containing Description and the first recognised claims section following it. This first claims section is treated as the original-language claims; check that it is the intended section in both inputs. Later translated claims sections are excluded. The comparison retains two section headings, which do not add to the word counts, and reports three sets of the same statistics:

| Scope | Text counted |
| --- | --- |
| Description | Description text only. |
| Full specification | Description plus the original-language claims. |
| Claims | Original-language claims only. |

Cover material, drawings, translated claims, and the reference bibliography are omitted from the EP comparison and its counts. Headers and footers are also excluded. Tables and inline graphics within retained sections stay in the comparison; text describing drawings inside Description remains description content. All three sets of statistics come from the same Word comparison. A passage that Word recognises as a move remains retained, including a move between Description and Claims; if Word instead records deletion and insertion, those inserted words contribute to the relevant scope.

The tool identifies sections automatically from standalone body-paragraph headings. It recognises `Description` or `Beschreibung`, followed by `Claims`, `Ansprüche`, `Patentansprüche`, or `Revendications`. Matching ignores letter case and permits a final colon. Headings inside tables, text boxes, headers, or footers are unsupported. Separate drawings and reference lists must have recognised standalone headings, such as `Drawings`, `Zeichnungen`, `Dessins`, or `References cited in the description`, to establish where they begin.

Both prepared specifications and uncropped publication files need valid boundaries; EP does not infer a section from its page number. The first claim must use `1.`, `1)`, `(1)`, or Word automatic numbering. If either input has missing or repeated section headings, ambiguous boundaries, empty sections, or an unrecognisable first claim, EP stops rather than returning guessed statistics. Prepare a copy containing Description and the numbered original-language Claims in order, with clear standalone headings, then retry. Check that OCR or PDF conversion has preserved the text and headings.

## PDFs and scanned pages

For a text PDF, Word converts the PDF into editable text before comparison. Conversion can alter line breaks, reading order, tables, and layout. Check columns, headers, footnotes, hyphenated line endings, and complex tables especially carefully.

For Full Document, images in the converted PDF conservatively trigger an OCR requirement so a scanned page is not silently excluded. A logo or diagram can also trigger this check. In EP, the check applies to images inside the retained Description and Claims sections; images in excluded cover or drawing pages do not by themselves require OCR. Select local OCR for the affected PDF or use an editable Word source, then check the converted text.

For a PDF with scanned pages, enable local OCR for that input. OCR processes the whole PDF, including any pages that already contain text, and produces an editable Word document. Recognised line positions reconstruct page sizes, approximate text sizes, indents and spacing. The recognised words, line order and page boundaries are preserved from the earlier OCR path. Missing or invalid layout information falls back to simple formatting, with a review note. Original font families, bold/italic styling, diagrams and semantic tables cannot be reliably recovered by Windows OCR; complex columns and tables can reflow. For close reproduction, use ABBYY separately to prepare and check a formatted DOCX, then compare that Word file. Recognition errors can appear as differences and affect crossover; check names, numbers, small print and mixed-language pages against the PDF.

OCR uses one installed recognition language for the document. It prefers UK English, then US English, then an available language from the Windows profile or installed language list. It does not detect a different language for each page or download language packs. Check that the selected language suits the document; ask IT to install required Windows OCR language support through approved channels.

OCR must confirm that every page was processed. If any page returns no recognised text, the comparison stops rather than reporting a figure that could omit scanned content. Check the source: use a verified local Word transcription, or remove genuinely blank pages from a copy before retrying. A page with only a picture can also trigger this check. OCR processes the whole supplied PDF in EP too; prepare a PDF containing the required sections or use a verified Word transcription if excluded drawing or blank pages prevent OCR from completing.

The built-in OCR path applies to PDFs. It does not recover text from images embedded in an ordinary Word input. Convert or transcribe those contents separately if they need to be compared.

## Confidentiality and local storage

The application is designed to work locally. It does not call an online comparison or OCR service and does not include application telemetry, a document history, or saved user settings. Word performs comparisons on the PC; scanned-PDF recognition uses Windows local OCR.

Archive and new documents can be selected from local folders, company OneDrive folders or accessible shared drives in File Explorer. UNC paths and mapped network drives are supported for source files and batch folders. The app copies selected files to its private local working folder and runs Word comparison and OCR on this PC. It does not modify the originals. Use a File Explorer path rather than a web sharing link; access to a shared drive requires the usual connection and read permissions.

An online-only file may need OneDrive to download it before the comparison can proceed. If it cannot be read, right-click the file in File Explorer, choose **Always keep on this device**, and wait for the download to complete before retrying. Offline comparison requires the files to be available on the PC. See [Microsoft's Files On-Demand guidance](https://support.microsoft.com/en-US/onedrive/save-disk-space-with-onedrive-files-on-demand-for-windows).

The default results folder and private temporary working files remain local. Results can also be saved to the actual Windows Desktop, another company OneDrive folder or an accessible shared-drive folder, including UNC paths and mapped drives. You need write permission, enough space and connectivity to the selected destination. Keep the drive connected throughout a batch; if it becomes unavailable, read the error/report status and restore access before retrying. Comparison documents, summaries and batch CSVs follow the destination's access, sync and backup rules. Private working folders must still be local and outside known OneDrive locations; shared-drive support applies to selected inputs and published results, while comparison and OCR stay on this PC.

This does **not** guarantee that the wider Office or Windows environment is offline. Microsoft 365 connected experiences, Office add-ins, endpoint monitoring, antivirus, backup, and cloud synchronisation are governed by the user's environment and IT policies. IT must configure or block those separately when required. Choosing a different output folder can put results under that folder's sync or backup rules even if the application cannot recognise them.

Word macros and automatic link updates are disabled for the comparison session. External relationships in Word XML working copies and the exported comparison are detached, retaining available cached content; linked features or hyperlink functionality may be lost. Fields use cached results rather than refreshed values. Legacy `.doc` files require Word to open the binary format before their contents can be processed, so Office's own protections remain essential. These measures do not replace an IT-enforced offline environment.

Inputs, converted documents, OCR text, job paths, and intermediate results may be written temporarily beneath `%TEMP%\LocalDocumentCompare`. The application attempts to remove its working directory after completion or cancellation and reports a leftover location if cleanup fails. A process or PC crash can leave working files behind. Cleanup covers the application's own working folder; it does not manage Office recovery locations or Windows caches. Temporary files are protected by the account's Windows permissions, not an encrypted application vault. Local deletion is not guaranteed secure erasure.

Saved comparison files and their summaries remain in the selected results folder until the user deletes them. The default results location is local application data to avoid automatically choosing a cloud-synchronised Documents folder. Follow your organisation's access, encryption, retention, backup, and disposal rules for both results and any residual temporary files.

## Limitations and troubleshooting

- **Language mismatch or uncertain language:** select archive and new text in the same source language. For Unknown/Mixed, inspect the actual relevant text before using **Language checked — compare**. Changing Word's proofing-language setting does not establish that the sources match; a known mismatch stays blocked.
- **Word does not start or the comparison stalls:** open Word normally, complete activation or first-run dialogs, and check for a hidden prompt. Follow local IT guidance for blocked automation or untrusted files. Cancellation can leave a failed job's working files, which should be checked if the tool reports their location.
- **Scanned text is missing:** use the local OCR option on the affected PDF and verify that an appropriate Windows OCR language is installed.
- **The layout changes:** Word's PDF conversion and especially OCR can change presentation. Use Word originals where available and assess text accuracy before relying on the percentage.
- **Changes look unexpected:** confirm that the files were selected in the correct archive/new order, review the original sources, and inspect PDF/OCR conversion quality. Word's matching of large reorganisations is not always equivalent to a reader's assessment.
- **Deletions seem missing:** switch Word to All Markup. The comparison retains revisions; its display is controlled by Word.
- **Protected, encrypted, or restricted files:** password-protected documents and documents requiring Protected View handling are unsupported. Use an authorised editable copy where policy permits. The tool does not bypass passwords, information protection, or organisational controls.
- **Complex document features:** equations, embedded objects, linked content, text in images, unusual fields, and heavily revised documents require manual validation. Check cached field values against the inputs. Visual changes are not a complete text comparison.

When reporting a problem, record the Windows and Word versions, input formats, whether OCR was selected, and the displayed error. Do not send confidential source files, screenshots, filenames, or document excerpts outside approved support channels.
