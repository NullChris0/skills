# 员工年度考核 Golden Fixture

This fixture freezes the confirmed stage-one count:

- `count/计数模型.json`: authoritative count model consumed by stage two
- `count/功能点计数证据.xlsx`: four-sheet human projection

Source baseline:

- backend `c761f4f529a80e2972436292f1aef35828fcf5a9`
- frontend `bbdb9f02580432a4fd9225e92e5351cf3ada3610`

Expected result: 11 ILF, 2 ELF, 22 EI, 10 EO, 5 EQ, 267 UFP.

`delivery/delivery-model.json` freezes the stage-two input for the generic builder.
`one-off/` preserves the employee-specific reference implementation used to
calibrate layout; it is not a runtime input and must not be copied into the
skill.
