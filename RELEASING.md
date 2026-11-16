# Publishing a new version (for Lisa)

Colleagues only receive versions you publish as a **Release** on GitHub. Saving or pushing code to the repository does nothing for them until you do the steps below, so it's safe to work on changes in between.

## One-time setup

1. Create a **public** repository on GitHub (for example `canvas-course-manager`).
2. Upload `canvas_course_manager.py`, the launcher files, `.gitignore`, `CHANGELOG.md`, and `RELEASING.md`. Upload `.gitignore` first so private files can't slip in.
3. In `canvas_course_manager.py`, set the repository name near the top and upload the file again:

   ```python
   GITHUB_REPO = "your-github-name/canvas-course-manager"
   ```

4. Check the repository page: you should **not** see `canvas_config.txt`, `course_settings.json`, any `.log` file, or any spreadsheet. If you ever do, delete it on GitHub right away and make a new Canvas access token (Canvas > Account > Settings), because the old one should be treated as public.

## Every new version

1. **Test first.** Open the program on your sandbox course, run a Preview, then a live run with APPLY. Make sure the new feature works and nothing old broke.
2. **Raise the version number** near the top of `canvas_course_manager.py`:
   - small fix: `1.0.0` → `1.0.1`
   - new feature: `1.0.1` → `1.1.0`
   - big change in how colleagues use it: `1.1.0` → `2.0.0`
3. **Add a CHANGELOG.md entry** at the top, in plain words for teachers.
4. Upload both files to GitHub.
5. On the repository page, click **Releases** → **Draft a new release**.
   - **Tag:** the version with a `v` in front, for example `v1.1.0`. It must match `__version__` exactly, or the program refuses to install it.
   - **Title:** for example `Version 1.1.0`.
   - **Description:** paste the CHANGELOG entry. Colleagues see this text in the update window.
6. Click **Publish release**. Colleagues are offered the update the next time they open the program.

## Good to know

- Your own copy doesn't nag you: if the program folder is a git checkout (has a `.git` folder), the automatic check is skipped. The **Check for updates** button still works.
- "Pre-release" and "Draft" releases are never offered to colleagues, so you can use them to test the update yourself.
- If an update goes wrong for someone, they can open the `backups` folder, copy the newest file out, rename it to `canvas_course_manager.py`, and put it back next to the launchers.
