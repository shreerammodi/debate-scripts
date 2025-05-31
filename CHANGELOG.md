# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to
[Semantic Versioning](https://semver.org/spec/v2.0.0.html).

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

[1.2.1]: https://github.com/shrimpram/debate-scripts/compare/v1.2.0...v1.2.1
[1.2.0]: https://github.com/shrimpram/debate-scripts/compare/v1.1.0...v1.2.0
[1.1.0]: https://github.com/shrimpram/debate-scripts/compare/v1.0.0...v1.1.0
[1.0.0]: https://github.com/shrimpram/debate-scripts/releases/tag/v1.0.0
