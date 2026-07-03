# Workspace Git Rules for Antigravity Agents

When performing work in this repository, you must adhere to the following Git branching and merging workflow:

1. **Preparation**:
   - Before starting any task, make sure you are on `main` and run `git fetch` and `git pull` to ensure you start from the newest remote version.

2. **Branching**:
   - Before making any code modifications or running import scripts, create a new local branch from `main`.
   - Name the branch descriptive of the task (e.g. `feat/amazon-import-support`, `fix/category-mapping`).

3. **Committing**:
   - Commit all code edits and script updates to your working branch with clear, atomic commit messages.

4. **Merging & Cleanup**:
   - Once all work is completed and successfully verified, merge your working branch back into the `main` branch.
   - Provide a concise title and description of your changes in the merge/commit message.
   - If the merge is successful, delete the working branch.
   - Run `pip freeze > requirements.txt` (or UV equivalent) to keep dependencies up-to-date, and commit any changes.
   - Finally, push the updated `main` branch to the remote repository.
