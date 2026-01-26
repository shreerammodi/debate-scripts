# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [3.2.0] - 2026-01-26

### Added

- Enable setting custom send and read doc folders
  ([44ab77760b](https://github.com/shreerammodi/debate-scripts/commit/44ab77760b))

## [3.1.1] - 2026-01-26

### Fixed

- Prevent Tags from being inadvertently deleted by filtering out style aliases
  ([b440367992](https://github.com/shreerammodi/debate-scripts/commit/b440367992))

## [3.1.0] - 2025-12-11

### Added

- GPL v3 License

### Changed

- Made duplicate doc closing logic more efficient
  ([451ceced4b](https://github.com/shreerammodi/debate-scripts/commit/451ceced4b))

## [3.0.0] - 2025-11-28

### Added

- Customization of document directories for send and read docs
  ([fd8e616080](https://github.com/shreerammodi/debate-scripts/commit/fd8e616080))

## [2.6.0] - 2025-11-28

### Added

- Acronym macros
  ([d71a415153](https://github.com/shreerammodi/debate-scripts/commit/d71a415153))

## [2.5.0] - 2025-11-19

### Added

- Customization of `ForReference` shrink size
  ([fb6d2c5a70](https://github.com/shreerammodi/debate-scripts/commit/fb6d2c5a70))

### Changed

- Delete all headers in `SendDocNoHeaders`
  ([d98b6df06d](https://github.com/shreerammodi/debate-scripts/commit/d98b6df06d))

## [2.4.1] - 2025-11-14

### Fixed

- Variable declaration in `Forreference`
  ([d357aa808b](https://github.com/shreerammodi/debate-scripts/commit/d357aa808b))

## [2.4.0] - 2025-11-14

### Added

- Let users customize `ForReference`
  ([9cb604e02b](https://github.com/shreerammodi/debate-scripts/commit/9cb604e02b))

### Changed

- Made helper functions public
  ([6e6040e0ee](https://github.com/shreerammodi/debate-scripts/commit/6e6040e0ee))

## [2.3.0] - 2025-10-26

### Added

- `SendDocNoHeadings` function
  ([6e5cc46d9e0e](https://github.com/shreerammodi/debate-scripts/commit/6e5cc46d9e0e))

### Changed

- Made zap public
  ([317f91148251](https://github.com/shreerammodi/debate-scripts/commit/317f91148251))
- Zapper now deletes undertags
  ([65d7c78e26f5](https://github.com/shreerammodi/debate-scripts/commit/65d7c78e26f5))

## [2.2.1] - 2025-10-17

### Changed

- Improved performance of zapper for Emory style citations
  ([8bb19a0d1b11](https://github.com/shreerammodi/debate-scripts/commit/8bb19a0d1b11))

## [2.2.0] - 2025-10-17

### Changed

- Implemented fast and slow versions of `ForReference`
  ([a4079aaf3497](https://github.com/shreerammodi/debate-scripts/commit/a4079aaf3497))

### Fixed

- Fixed zapper behavior for Emory citation
  ([0a83ca47811a](https://github.com/shreerammodi/debate-scripts/commit/0a83ca47811a))

## [2.1.0] - 2025-10-06

### Changed

- Combined ForReference and HighlightToFill
  ([45e0ea2cb9](https://github.com/shreerammodi/debate-scripts/commit/45e0ea2cb9))
- Use VBA preset gray color for
  ForReference([1b74062852](https://github.com/shreerammodi/debate-scripts/commit/1b74062852))

### Fixed

- Enable multiple read docs to be open
  (<https://github.com/shreerammodi/debate-scripts/issues/7>)
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

[3.2.0]: https://github.com/shreerammodi/debate-scripts/compare/v3.1.1...v3.2.0
[3.1.1]: https://github.com/shreerammodi/debate-scripts/compare/v3.1.0...v3.1.1
[3.1.0]: https://github.com/shreerammodi/debate-scripts/compare/v3.0.0...v3.1.0
[3.0.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.6.0...v3.0.0
[2.6.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.5.0...v2.6.0
[2.5.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.4.1...v2.5.0
[2.4.1]: https://github.com/shreerammodi/debate-scripts/compare/v2.4.0...v2.4.1
[2.4.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.3.1...v2.4.0
[2.3.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.2.1...v2.3.0
[2.2.1]: https://github.com/shreerammodi/debate-scripts/compare/v2.2.0...v2.2.1
[2.2.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.1.0...v2.2.0
[2.1.0]: https://github.com/shreerammodi/debate-scripts/compare/v2.0.0...v2.1.0
[2.0.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.2...v2.0.0
[1.2.2]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.1...v1.2.2
[1.2.1]: https://github.com/shreerammodi/debate-scripts/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/shreerammodi/debate-scripts/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/shreerammodi/debate-scripts/releases/tag/v1.0.0
