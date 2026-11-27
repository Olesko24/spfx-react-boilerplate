# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/).

## [4.0.0] - 27.11.2026
### Changed
- Upgraded SPFx Version to 1.24.2 with the new Heft toolchain (replaces Gulp)
- Upgraded React to 18.3.1 (`createRoot`)
- Upgraded Node.js Version to v24.14.1, TypeScript to 5.8, ESLint 9 (flat config)
- Upgraded TailwindCSS to 4.3.5, Tailwind now runs automatically in `npm start` / `npm run build`
- Upgraded deployment script to @pnp v4
- `initSP` no longer calls SharePoint (previously could create the Site Assets library on load)
- Unknown hosts fall back to SharePoint mode instead of throwing

### Fixed
- `ensureList` crashed because options were spread as arguments
- `addViewToList` crashed without options and sent `viewXml` as view property
- `getListItems` paged beyond 5000 items, loads attachments in one request and expands lookups without `fields`
- `getListFields` no longer swallows errors
- Tailwind `@source` path

### Removed
- Gulp, `spfx-fast-serve`, unused `@microsoft/sp-adaptive-card-extension-base`
- Unused SCSS module, assets, strings and props
- Outdated README sections (husky/commitlint, Gulp issues)

## [3.1.0] - 12.08.2025
### Changed
- Upgraded SPFx Version to 1.21.1
- Upgrade Node.js Version to v22.17.1
- Upgraded TailwindCSS Version to 4.1.11
- Tailwind Docs

### Added
- Fluent UI / SharePoint Theme Examples

## [3.0.0] - 08.10.2024
### Added
- spfx-fast-serve for better & faster development [Link](https://github.com/s-KaiNet/spfx-fast-serve)

### Changed
- New SPFx Version 1.19 with Node Support for v.18
- New @pnp/sp Version 4.5.0

## [2.0.0] - 10.10.2023
### Changed
- TailwindCSS is now default on main branch

## [1.0.0] - 19.06.2023
### Initial
- Initial Relase

### Added
- TailwindCSS Support