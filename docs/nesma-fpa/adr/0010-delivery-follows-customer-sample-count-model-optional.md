# Delivery follows the customer sample; the count model becomes an optional internal check

Supersedes ADR 0009.

The customer rejected deliverables built from the count model: a ten-chapter document with a "data model" chapter, interface-contract card images, and NESMA identifiers and FP in the feature list. Their evaluators compare the feature list and the requirement document line by line against their own sample, and cannot search fields inside images. Three systems were then rewritten directly from code and existing business material to match the sample and were accepted.

Stage two now takes the customer's sample as the authority for structure and wording. Its inputs are the estimated system's code and the user's business material. The feature list is settled with the user first; chapter 8 of the requirement document is generated from it and must match it row for row.

`计数模型.json` is no longer required. When present, it is used only to check that every data function and transaction maps to a feature-list row, through an internal mapping file that is never delivered. Delivered files carry no NESMA identifiers, FP, UFP, or person-months.

Word is produced by copying prototype paragraphs from a layout template taken from an accepted deliverable, then refreshing the table of contents with LibreOffice. The C# OpenXML generator, `delivery-model.json`, and interface-contract cards are removed. Physical-table cards remain as an optional appendix.

Old deliverables may now supply facts and images, but not structure. Overwriting an existing deliverable is allowed only on the user's explicit request, after backing it up.
