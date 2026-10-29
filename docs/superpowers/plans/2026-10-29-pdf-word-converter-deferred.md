# Deferred findings — PDF Converter v1 (2026-10-29)

Findings the per-task and final reviews rated Minor or parked with a ruling. None blocks v1; they are the backlog for v1.1.

- Task 1: minor (deferred): pytest.ini filterwarnings for pymupdf typing.ByteString DeprecationWarning may go stale on pymupdf upgrade
- Task 1: minor (deferred): fixtures mixed_pdf, corrupt_pdf, sample_docx have no dedicated test (exercised by Tasks 3, 5)
- Task 2: minor (deferred): no tests for ValueError on unknown mode or unique_path on suffix-less source
- Task 2: minor (deferred): unique_path has check-then-use race (sequential worker + .part rename makes it harmless)
- Task 3: minor (deferred): per-page pymupdf exceptions inside inspect_pdf's loop escape unwrapped (Task 4's pdf_to_docx maps any exception to MSG_FAILED, so the user still sees the right text)
- Task 3: minor (deferred): no tests for exact-half scanned boundary or 0-page PDF
- Task 4: minor (deferred): no test for MSG_FAILED branch / part cleanup on mid-conversion failure
- Task 4: minor (deferred): cv.close() in finally runs before logging restore; a raise there would skip the restore — nest try/finally
- Task 4: minor (deferred): exception inside progress callback propagates into pdf2docx; progress handler on root logger not thread-filtered
- Task 4: minor (deferred): original test_root_logger_level_is_restored still vacuous (superseded by progress variants); failure variant does not assert MSG_CANNOT_SAVE
- Task 5: concern (open): legacy .doc encryption/write-protection pre-check untested on real files
- Task 5: minor (deferred): comment claims wrong password makes Word raise (false on Word 16); faulthandler fixture restore should pass explicit file=; Quit() without SaveChanges=0; Open without ConfirmConversions=False; TemporaryDirectory without ignore_cleanup_errors; LibreOffice path skips pre-check and lacks -env:UserInstallation; regex assumes w: prefix; no tests for MSG_CANNOT_SAVE / LibreOffice timeout
- Task 5: minor (deferred): FIB bit 0x04 (fReadOnlyRecommended) not refused — unverified whether it prompts; _has_write_password returns False on BadZipFile/OSError (corrupt zip reaches Word); macro disabling has no automated test
- Task 6: minor (deferred): _current set before output assigned (harmless window); add() notify/put outside lock could reorder ids under concurrent adds; no tests for on_update raising, generic-exception path, docx_to_pdf dispatch; stop test lacks try/finally around release
- Task 6: minor (deferred): stopped job gets no final on_update (moot: stop() only runs while the window is closing); test waits on private w._current
- Task 7: minor (deferred): selector built from job.id string (fine for ints); list does not scroll independently (footer moves off-screen with many jobs); no :focus-visible styles; no color-scheme for dark scrollbars; role=tablist without tabpanel/arrow keys; 5px "Waiting" alignment doubt in s1 (not reproduced in s3)
- Task 8: minor (deferred): open_folder/show_file/int(job_id) unguarded (pywebview logs, nothing on screen); drop_paths assumes dict entries; on_shown swallows all exceptions; no Api unit tests / no test that Api's public names are exactly the five methods
- Task 9: minor (deferred): README "Known limits" names only password-protected PDFs (Word files are refused too); omits never-overwrite naming, one-at-a-time, banner, log path, exe size; splash uses "..." not "…"; icon has opaque white background
- Final: parked — in-flight push inside evaluate_js when closing starts can block, making main() wait up to ~190 s before exit (narrow race, nothing orphaned) — Ruling: real and deferred; fix candidate: evaluate_js with callback or skip the join when no Word PID was ever found
- Final: parked — watchdog disabled (with a warning) if Word's window is not found within 5 s; pre-timer COM calls (DispatchEx, property sets) unwatched — Ruling: real and deferred; cost: a hang only in an environment where Word has no "OpusApp" window
- Final: parked — closing during a PDF→Word job does not wait; a `.<pid>.part` can remain if the daemon thread dies mid-save — Ruling: pre-existing, deferred
- Final: parked — startup-failure dialog names a log that may not exist when logging setup itself failed; webview.start() exception after show reported as "could not start" — Ruling: cosmetic, deferred
- Final: parked — rename collision yields `Report (1) (1).docx` instead of `Report (2).docx` — Ruling: acceptable (race only, never overwrites)
- Final: parked — LibreOffice timeout kills soffice.exe only (soffice.bin may survive); fresh profile adds first-run time — Ruling: deferred, LibreOffice not on dev machine
- Font substitution: table key SCHOLA also matches a hypothetical font named "Scholarly" (maps to Century Schoolbook instead of the class fallback) — harmless degrade; pdf2docx still glues words on some tightly justified LaTeX lines (engine limitation, not fixable by post-processing)
- Beta report (2026-10-29): fixed — words glued at PDF line ends next to non-ASCII characters („ ” ă î …) and LaTeX line-end hyphens left in ("pre-sentation"). Trade-off accepted: a real compound hyphen that falls at a line end before a lowercase word is removed too ("post-credits" → "postcredits")
- Beta report (2026-10-29): open — pdf2docx lays out each PDF page as its own section with exact line spacing and footer text pushed down by spacing; when Word wraps a paragraph onto one more line than the PDF (substituted font), the page number spills onto an extra, otherwise blank page
