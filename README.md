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

## Daily Device-Switch Workflow

### Start Work On Any Computer

Always pull the latest version before editing files:

```powershell
cd D:\codex\teaching-workspace
git config --global http.https://github.com.proxy http://127.0.0.1:7890
```

```powershell
cd D:\codex\teaching-workspace
git pull
```

If Git says `Already up to date.`, this computer already has the newest files.

### Before Switching To Another Computer

Save your local changes to GitHub:

```powershell
cd D:\codex\teaching-workspace
git status
git add .
git commit -m "Update teaching materials"
git push
```

Use a clearer commit message when possible, for example:

```powershell
git commit -m "Update lesson 01 slides"
git commit -m "Revise lesson 02 lesson plan"
git commit -m "Add animation workflow references"
```

### If There Is Nothing To Commit

If this command:

```powershell
git commit -m "Update teaching materials"
```

prints `nothing to commit, working tree clean`, it means Git did not detect any new changes. This is not an error.

You can then run:

```powershell
git status
```

If the working tree is clean, there is nothing else to upload.

### Golden Rule For PowerPoint And Word Files

Do not edit the same `.pptx` or `.docx` file on two computers at the same time.

Recommended habit:

1. Before editing, run `git pull`.
2. Finish editing on one computer.
3. Run `git add .`, `git commit`, and `git push`.
4. Only then continue on the other computer.

PowerPoint and Word files are binary files, so Git cannot merge two different edits as easily as it can merge text files.

### If Git Reports A Conflict

Stop and inspect the message before continuing. Do not randomly delete files or force a push.

For binary files like `.pptx` or `.docx`, usually you must choose one version to keep.

Useful first command:

```powershell
git status
```

Then decide which file version is correct before committing again.

## Structure

```text
lesson-plan/   Course lesson docs, decks, notes, references, and scripts
powerpoint/    Reusable templates and PowerPoint projects
archive/       Old versions and reference copies kept out of the main workflow
```

Generated preview folders such as `ppt_preview` should not be committed because they can be regenerated.
