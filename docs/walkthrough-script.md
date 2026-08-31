# Transfer News Generator — presenter script

This script follows the 56-second narrated and captioned video on the **How to use** page. The video contains sanitized examples only.

## 0:00–0:08 — Welcome

“Welcome to the Nigeria Uyo Mission Transfer System, developed by ThankGod Andrew. One verified web run creates every essential output.”

## 0:08–0:16 — Uploads

“Upload the current Transfer Management PDF, previous Transfer News PDF, current roster report, and previous roster report.”

## 0:16–0:24 — Publication details

“Confirm the mission, president, preparer, and detected transfer title. Override the title only when necessary.”

## 0:24–0:32 — Optional controls

“Keep the updated roster and statistics selected. A Word template and corrections CSV are optional.”

## 0:32–0:40 — Generate

“Generate the package and keep the tab open while validation, matching, verification, rendering, and ZIP creation finish.”

## 0:40–0:48 — Review

“Download the ZIP and inspect Verification, Changes, notes, and warnings before publishing. Resolve any blocking row first.”

## 0:48–0:56 — Privacy

“Download promptly, protect confidential records, and never commit real transfer files to GitHub.”

## Supplemental movement-plan walkthrough

Use this after the Transfer News video when training mission-office staff:

1. Open **Movement plan** in the main navigation.
2. Upload the previous and current Transfer News PDFs. If either table cannot be parsed, paste its `Missionary | Zone | Area` text in the fallback panel.
3. Upload the area-to-apartment directory. CSV or XLSX is recommended because staff can maintain it in Excel; JSON and the legacy Word document are also supported.
4. Add only confirmed historical exceptions in the format `Elder Twum: Anua Obio -> Mbierebe`.
5. Generate the plan and verify the comparison audit. New arrivals and released missionaries are skipped automatically; every movement is grouped under the previous zone.
6. Download the Word plan, review CSV, and editable apartment-directory CSV. The STATUS cells remain blank so drivers can tick them after each successful movement.
