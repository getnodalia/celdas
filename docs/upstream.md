# Celdas upstream

| Field | Value |
|---|---|
| Upstream | `storytold/gridcraft` |
| Last observed | `c0ee120ac2d9b59630f5abf1151dd9b471fe2926` |
| Last evaluated | `a36e4987c1b5b53af69a3f3573418e18612d1a4e` |
| Last integrated | `bea2191b354270d06048d1a6cbc14b5690ea0f40` (282 commits since the last sync) on 2026-11-18 |
| Skipped | the community button and the brand files (Celdas has neither) |
| Adapted | every name (`scripts/rebrand.py`), the app icon, the Nodalia interface |
| Conflicting | `README.md` (kept ours), `AGENTS.md` (upstream plus our quality contract) |

## How the sync works

Upstream's code arrives renamed: `scripts/rebrand.py` applies `.github/rebrand.toml` (GridCraft
becomes Celdas, internal names included; links to the shared libraries point at
getnodalia/foundation) to upstream's tree, and the weekly sync
(`.github/workflows/upstream-sync.yml`) records the result as one mirror commit: its only parent is
the previous mirror commit, its subject is `Upstream <sha> with this app's names`, and an
`Upstream-Commit:` trailer names the upstream commit it renames. The sync merges that commit into a
branch and opens a pull request; files listed in `.github/upstream-keep` keep this repository's
version. No mirror branch is pushed, so merge sync pull requests with a merge commit: a squash would
drop the mirror commit the next sync starts from.

The mirror starts at the renamed `a36e4987`, merged into `main` (with this repository's tree) by the
pull request that renamed the app. `python3 scripts/rebrand.py check` fails when an old name comes
back; a line that must keep one says `rebrand: keep`.