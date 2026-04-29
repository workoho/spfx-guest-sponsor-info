<!-- markdownlint-disable MD041 -->
## Description

<!-- What does this PR change and why? Link related issues with "Fixes #123" or "Closes #123". -->

## Type of change

<!-- Check all that apply -->

- [ ] Bug fix
- [ ] New feature
- [ ] Performance improvement
- [ ] Refactoring (no behaviour change)
- [ ] Documentation
- [ ] Dependency update
- [ ] Infrastructure / tooling

## Checklist

- [ ] `npm run fix` and `npm run lint` pass with no errors
- [ ] Tests added or updated (if logic, components, or services changed)
- [ ] `npm test` passes
- [ ] New locale keys added to **all 17** locale files under `src/webparts/guestSponsorInfo/loc/`
- [ ] No direct Graph calls added — all data fetching goes through the Azure Function
- [ ] No `@microsoft/sp-*` or `react` versions changed individually
- [ ] Commit messages follow [Conventional Commits](https://www.conventionalcommits.org/) (enforced by pre-commit hook)
