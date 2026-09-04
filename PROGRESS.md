# TGS-2022015539 — Build Progress

## Power Platform environment (DONE)
- Name: `TGS-2022015539-Applications Integration with Power Apps and Power Automate`
- Id: `087172fe-d0cf-e38a-a9c3-106c04330a95`  Region: Asia  Type: Sandbox  Dataverse: Yes

## Flows in that environment (8, all named "Lab N - ... (DO NOT DELETE)")
Lab 3 Trigger and Actions · Lab 4 Log to Excel · Lab 5 Conditions and Branching ·
Lab 5b Daily Digest · Lab 6 Leave Application Approval · Lab 11 Submit Leave Request ·
Lab 12 Return Leave Balance · Lab 13 Approval Round Trip

## Canvas apps in that environment
Lab 8 - Leave Request App (DO NOT DELETE)

## Artifacts
- Deck v8.0 — 108 slides, 0 shapes past the slide bottom, 0 truncated strings IN THE
  RENDERED PDF (verified with tools/check_deck_render.py — the PPTX XML can be clean while
  the renderer clips, so the PDF text layer is the check that counts)
- LP v8.0 — 2-day schedule, 480 min/day excluding lunch
  (Day 1 = 480 delivery; Day 2 = 320 delivery + 160 assessment, ending 18:30)
- LG v8.0 — 35 pages, 14 lab screenshots embedded, steps restart at 1 per lab
- Assessment v4 — WA 5 questions (K1–K5) + PP 4 tasks (A1–A8), mirroring the TMS original
- 14 lab folders, each with its lab sheet, data workbook and flow packages

## QA fixes applied after the first audit
1. Deck flow-chip truncation — stage labels are now AUTHORED short in data_domainN.py
   ("stages"), never derived by truncating a step sentence. Build-time assertions in
   process_map now FAIL the build if a label >44 or caption >26 chars would be cut.
2. LP Day 2 overran to 9.7h — rebuilt to exactly 480 min (320 delivery + 160 assessment, the 10-minute assessment briefing counted as assessment time),
   ending 18:30. An assertion enforces 480 for every day.
3. LG step numbering ran 1→126 across all labs — each steps block now gets its own
   w:abstractNum + w:num, so every lab restarts at 1.
4. LG had no workflow screenshots — 14 unique screenshots captured from the live course
   environment, one per lab, embedded with captions.
5. Both TOCs were static text — inject_toc.py now emits a REAL Word TOC field
   (fldChar begin/separate/end + TOC instruction) whose cached result is the
   page-numbered entry list, so LibreOffice renders it AND Word can refresh it.

## Fixes applied after the second audit
6. Deck truncation RETURNED at render time (55 strings / 15 slides) even though the PPTX
   XML was clean: process_map chip captions sat in a 0.38in-high box, so a wrapped second
   line was clipped by the renderer. Boxes now have two lines of height, and
   tools/check_deck_render.py scans the RENDERED PDF — the only check that catches this.
7. Two lab screenshots were error pages (Lab 9 "there's been a disconnect"; Lab 14 four
   red failed-import banners). Both re-captured from working pages, and the capture step
   now asserts on page text before saving.
8. LP Day 2 gave Labs 13-14 only 10 minutes. Rebalanced: Topic 3 45 min, Topic 4 35 min,
   Labs 13-14 50 min. Still exactly 480.

## Guards added so these cannot regress
- build_slides.py asserts process_map labels <=44 and captions <=40 chars.
- build_lesson_plan.py asserts each day totals exactly 480 minutes.
- tools/check_deck_render.py — no "..." in the rendered deck PDF.
- tools/check_lab_screenshots.py — one screenshot per lab, perceptually distinct.

## Package formats
Legacy .zip (portable) + Dataverse solution .zip (imports on this tenant, verified).
The tenant has "Create in Dataverse solutions" enabled, which disables legacy import.
