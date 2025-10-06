# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [2.1.0] - 2025-10-06

### Changed

- Combined ForReference and HighlightToFill
  ([45e0ea2cb9](https://github.com/shreerammodi/debate-scripts/commit/45e0ea2cb9))
- Use VBA preset gray color for
  ForReference([1b74062852](https://github.com/shreerammodi/debate-scripts/commit/1b74062852))

### Fixed

- Enable multiple read docs to be open
  ([840199b0e0](https://github.com/shreerammodi/debate-scripts/commit/840199b0e0))

## [2.0.0] - 2025-09-21

### Added

- **Granular zapper**: enables zapping of a single card in the doc
  ([#3](https://github.com/shreerammodi/debate-scripts/issues/3))
  ([f827dcff9c](https://github.com/shreerammodi/debate-scripts/commit/f827dcff9cc50b0f6ab06858485e03d673cf39bc))

### Changed

- **Michigan compatibility**: ensures zapper and send-doc work with Michigan
  "Analytics" style
  ([f8eefa299c](https://github.com/shreerammodi/debate-scripts/commit/f8eefa299c565d239ca17550e87440484509305b))

### Fixed

- Check for styles before running scripts to prevent errors
  ([f85eafd585](https://github.com/shreerammodi/debate-scripts/commit/f85eafd5854c49d1e653d9112386ade5f3f1a5fb))

## [1.2.2] - 2025-06-01

### Fixed

- Prevent overriding user's selected highlight color

## [1.2.1] - 2025-05-31

### Added

- Update README with Mark macros from 1.2.0

### Fixed

- Properly normalize marker in `MarkCard`

## [1.2.0] - 2025-05-31

### Added

- Implement `MarkCard` macro which marks a card at the cursor.
- Implement `CompileMarkedCards` macro which creates a compiled list of all the
  marked cards in the document.

### Removed

- Excess documentation.
- Unused variables in zap-doc.

## [1.1.0] - 2025-05-19

### Changed

- Re-implement CondenseZap logic so it doesn't leave extra spaces in cards and
  processes documents faster.

## [1.0.0] - 2025-05-18

Initial Release

[2.1.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.0.0...v2.0.0
[2.0.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.2...v2.0.0
[1.2.2]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/shreerammodi/debate-scripts/releases/tag/v1.0.0
