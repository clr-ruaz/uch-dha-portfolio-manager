# Workspace setup checklist

- [x] Verify that `copilot-instructions.md` exists.
- [x] Clarify project requirements: SPFx 1.21.1, React 17, Single Part App Page.
- [x] Scaffold the project from the existing working DHA Portfolio Manager source.
- [x] Customize the manifest and package for standard and full-page hosts.
- [x] Install required extensions: none required.
- [x] Install dependencies and compile the project.
- [x] Create and run build/package tasks where needed.
- [x] Launch/debug skipped because it was not requested; deployment instructions are in README.
- [x] Verify README and final documentation.

## Project conventions

- Target SPFx version: 1.21.1.
- The web part must support `SharePointWebPart` and `SharePointFullPage` and set `supportsFullBleed` to `true`.
- Keep generated folders and package artifacts out of source control.
- Validate with `npm run build` and `npm run package-solution`.
