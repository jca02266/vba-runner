---
name: publish
description: Publish the repository artifacts by following the project-owned .claude/commands/publish.md procedure, including changelog/version review, build, authentication, publish, commit, tag, and push.
---

# Project publish workflow

Use `.claude/commands/publish.md` as the single source of truth. Read it in full
before acting and follow its artifact-selection, changelog, version-confirmation,
build, publish, and git release steps. Do not copy or restate that command's
artifact table, version rules, credentials, or commands here.

Before publishing, preserve unrelated user changes and verify the repository is
clean and the required test gate from the command has passed. Ask the user for
the version bump whenever the command requires confirmation; never infer a
publish version. Do not edit `.claude/commands/publish.md` as part of this skill.
