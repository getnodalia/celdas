# Nodalia migration: Celdas

Category: secondary app.
Status: renamed to Celdas throughout, own release workflows, as of 2026-11-13. Evidence: the pull
request that adds this file.

## Current state

Pure-Rust spreadsheet (egui), MIT OR Apache-2.0. Cargo git dependencies come from
getnodalia/foundation and nodalia-design. The gate, FreeBSD, Nix, packaging lint, AGENTS.md and
git-sources checks call getnodalia/foundation's reusable workflows pinned by commit. Releases run
on Celdas's own workflows (`release-auto.yml` with Foundation's release-tag and release-publish,
then `release-linux.yml`, `release-macos.yml` and `release-windows.yml`). Every name is Celdas's,
internal ones included. There is no upstream sync.

Celdas keeps `.xlsx` as its workbook format: upstream never had a file format of its own, so
there is no old extension to rename.

## Inherited dependencies

Resolved: fonts, translations and the shared libraries come from getnodalia/foundation's vendor branches, and nothing in `Cargo.toml` names another organisation.

## Technical debt

- Windows ARM64 is not packaged yet (the inherited workflow was removed); `release-windows.yml`
  builds x64.

## Reusable Nodalia components

getnodalia/foundation (vendored crates, fonts, reusable workflows);
getnodalia/nodalia-design (theme, icons, `nodalia-update`).

## NDL work pending

The interface in the Nodalia Design Language: islands, the hamburger menu with About,
Preferences and Keyboard shortcuts, NDL window controls on Linux only, the global menu over D-Bus.

## Recommended order

1. The NDL interface.
2. Windows ARM64 packages in `release-windows.yml`.
