# Publishing to the Public Repository

This project uses a dual-repo setup: a private repo with full history and a
public repo with squashed commits.

## Remotes

| Remote    | Repository                          | Branch           |
|-----------|-------------------------------------|------------------|
| `private` | `xlport/xlport-internal`            | `master`         |
| `public`  | `mkirsten/xlport`                   | `main`           |

Local branch `public-release` is an orphan branch (no shared history with
`master`) that maps to the public repo's `main`.

## Files excluded from the public repo

These files exist on `master` but are not on `public-release`:

- `Dockerfile`
- `build_and_upload.sh`
- `deploy.sh`
- `kube-*.yaml`
- `.env`
- `PUBLISHING.md` (this file)

## Publishing updates

```bash
# 1. Make sure master is up to date
git checkout master

# 2. Switch to the public-release branch
git checkout public-release

# 3. Squash-merge all new changes from master
git merge --squash master

# 4. Commit with a version/description
git commit -m "v2.1.0 - description of changes"

# 5. Push to the public repo
git push public public-release:main

# 6. Switch back to master
git checkout master
```

## Accepting contributions from the public repo

If someone submits a PR to the public repo:

```bash
# Merge the PR on GitHub, then pull it locally
git checkout public-release
git pull public main

# Cherry-pick or merge changes into master if desired
git checkout master
git cherry-pick <commit-hash>
```

## Initial setup reference

The public repo was created with an orphan branch (no history):

```bash
git remote rename origin private
git remote add public git@github.com:mkirsten/xlport.git
git checkout --orphan public-release
# (removed private files from staging)
git commit -m "Initial open-source release of xlPort"
git push public public-release:main
```
