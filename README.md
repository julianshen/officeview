# officeview

React components that render Office documents (Word `.docx`, PowerPoint `.pptx`, Excel `.xlsx`) on an HTML `<canvas>` — pixel-faithful rendering, **not** translation into DOM/HTML. Designed to work on both desktop and mobile browsers.

## Status

Project bootstrap in progress. Scaffold, tooling, and dependency installs are set up; component modules follow.

## Stack

- TypeScript (strict), React 18+/19
- Vite (lib-mode build) + Vitest (jsdom) for tests
- Rendering surface: `<canvas>` 2D context, with devicePixelRatio-aware scaling for crisp output on retina screens and mobile
