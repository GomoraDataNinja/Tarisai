Tarisai

Reconciliation bot for Supplier vs Creditors Ledger

Tarisai lets you upload a Supplier file, a Ledger file, and an Output Template.
It detects the main tables automatically, maps columns, matches transactions, and generates a reconciliation workbook you can download.

What Tarisai does

Reads multi-sheet Excel files

Detects the best table inside each sheet

Normalizes dates, references, and amounts

Matches transactions using the same practical logic an analyst would use

Document ID matching when available

Invoice reference matching

Fallback matching using amount and date window for payments

Produces a ready-to-download Excel output file

Includes match details, missing items, mismatches, and items that need review

Files you upload

Supplier file

Either a Supplier Statement or an Invoice List

Can include multiple sheets

Ledger file

Your creditors ledger extract

Can include multiple sheets

Template file

Your reconciliation output template

Tarisai writes results into the template and adds supporting tabs

Matching logic used

Tarisai tries matches in this order:

DocID totals match

If both sides contain DocIDs like HREINV or HRECRN

Groups by DocID and compares totals within tolerance

Invoice key match

Extracts invoice references from ledger “external document” fields

Matches supplier invoice totals to ledger totals

Payment fallback

If an invoice reference is missing, it tries amount + date window matching

Uses text token overlap as a confidence booster

Each match includes:

Match score (confidence)

Match reason (plain explanation)

Status (Matched, Needs review, Amount mismatch, Missing on Supplier, Missing in Ledger)

Output workbook tabs

Main Template sheet populated with reconciliation tables

Match_Detail

missing_in_ledger

missing_on_supplier

amount_mismatch

needs_review
