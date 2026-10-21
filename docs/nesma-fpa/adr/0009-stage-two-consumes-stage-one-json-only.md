# Stage two consumes only the stage-one count model

Stage one owns functional size. Stage two owns delivery structure and interface rendering. The only contract between them is `计数模型.json`; the workbook is a projection for human review.

Stage two must not read old delivery files or an old `api-tables` directory for facts, IDs, types, FP, or interface contracts. Its interface YAML schema is internal to the delivery skill. Each run writes a new delivery directory with Word, feature XLSX, mapping, YAML, card sources, and images. Delivery artifacts are never written back to the estimated system's repository.
