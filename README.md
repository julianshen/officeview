# officeview

React components that render Office documents (Word `.docx`, PowerPoint `.pptx`, Excel `.xlsx`) on an HTML `<canvas>` — pixel-faithful rendering, **not** translation into DOM/HTML. Designed for both desktop and mobile browsers.

## Status

Project bootstrap complete: tooling scaffold, dependency installs, test environment, and lib-mode build all verified. Component modules follow next.

## Stack

- **TypeScript** (strict) + **React** 18/19 (peer dependency)
- **Vite** (lib-mode build) for packaging
- **Vitest** + jsdom + `canvas` package for tests (real Canvas 2D in Node)
- Rendering surface: `<canvas>` 2D context with devicePixelRatio-aware scaling for crisp output on retina and mobile screens

## Scripts

| Script | Purpose |
| --- | --- |
| `bun run build` | Type-check + lib build to `dist/` |
| `bun run dev` | Vite dev server (for demos/sandbox) |
| `bun run test` / `bun run test:watch` | Vitest (jsdom) test suites |
| `bun run lint` | `tsc --noEmit` type check |

## License

MIT
