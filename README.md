# Teaching Workspace Sync

This private repository syncs teaching materials between computers.

- `lesson-plan`
- `powerpoint`
- `archive`

Large Office and image assets are configured for Git LFS. Install Git LFS before the first commit or clone.

## Repository Location

Use the same local path on each computer:

```text
D:\codex\teaching-workspace
```

## Clone On Another Computer

```powershell
git lfs install
git clone https://github.com/morille/teaching-workspace.git D:\codex\teaching-workspace
```

## Daily Workflow

Start work:

```powershell
cd D:\codex\teaching-workspace
git pull
```

Before switching computers:

```powershell
cd D:\codex\teaching-workspace
git status
git add .
git commit -m "Update teaching materials"
git push
```

## Structure

```text
lesson-plan/   Course lesson docs, decks, notes, references, and scripts
powerpoint/    Reusable templates and PowerPoint projects
archive/       Old versions and reference copies kept out of the main workflow
```

Generated preview folders such as `ppt_preview` should not be committed because they can be regenerated.
