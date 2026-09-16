# CLAUDE.md

## Git workflow

- Before starting work and before every push: `git pull --rebase origin main`. Dave and Jared both work here.
- When a unit of work is done and checks pass, commit it locally. Small, descriptive commits.
- Review the diff before every commit for secrets, keys and tenant IDs. If one was already committed, stop and report it.
- Client names and contact details are expected in a repo that exists for that client's engagement. Flag them only when they would land in a public repo or in a repo not specific to that client.
- Pushing to `main` deploys to production (Railway). Push once at the end of a working session, or when asked for a deploy. Run the build and tests locally first, never push work in progress, and state in your summary that the push triggered a deploy.
