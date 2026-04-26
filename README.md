# Teaching Workspace Sync

This private repository is intended to sync two working folders between computers:

- `lesson plan`
- `powerpoint`

Large Office and image assets are configured for Git LFS. Install Git LFS before the first commit or clone.

## First Setup On This Computer

```powershell
git init
git lfs install
git add .gitattributes .gitignore README.md "lesson plan" "powerpoint"
git commit -m "Initial teaching workspace sync"
git branch -M main
git remote add origin https://github.com/morille/YOUR-PRIVATE-REPO.git
git push -u origin main
```

## Clone On Another Computer

```powershell
git lfs install
git clone https://github.com/morille/YOUR-PRIVATE-REPO.git
```

## Daily Workflow

Before switching computers:

```powershell
git status
git add "lesson plan" "powerpoint"
git commit -m "Update teaching materials"
git push
```

After switching computers:

```powershell
git pull
```

The `lesson plan/ppt_preview` folder is ignored because it can be regenerated from the PowerPoint file.
