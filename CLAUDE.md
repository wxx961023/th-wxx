# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Development Commands

### Core Development
- `pnpm dev` - Start development server with increased memory allocation
- `pnpm build` - Production build with memory allocation and cleanup
- `pnpm build:staging` - Staging environment build
- `pnpm build:static` - Static build mode
- `pnpm preview` - Preview production build locally
- `pnpm typecheck` - Run TypeScript and Vue type checking

### Maintenance
- `pnpm clean:cache` - Clean all caches and reinstall dependencies
- `pnpm svgo` - Optimize SVG files
- `pnpm report` - Build with bundle analysis

**Note:** All linting tools (ESLint, Prettier, Stylelint) are currently disabled as per package.json

## Architecture Overview

This is a Vue 3 + TypeScript admin template based on `pure-admin-thin`, a simplified version of `vue-pure-admin`. The project is configured for bill download and PDF processing functionality.

### Key Technologies
- **Vue 3** with Composition API and `<script setup>`
- **TypeScript** with relaxed strict mode (`strict: false`)
- **Element Plus** for UI components
- **Pinia** for state management
- **Vue Router 4** with simplified auth
- **Vite** for build tooling
- **Tailwind CSS** for styling
- **@pureadmin** suite of components

### Project Structure
- `src/views/` - Page components (welcome, login, error pages)
- `src/layout/` - Layout components and navigation
- `src/components/` - Reusable components (ReAuth, ReDialog, ReIcon, etc.)
- `src/router/` - Routing configuration with simplified static routes
- `src/store/` - Pinia store modules (user, app, settings, etc.)
- `src/utils/` - Utility functions and helpers
- `src/api/` - API service definitions
- `src/directives/` - Custom Vue directives
- `src/plugins/` - Plugin configurations

### Current Application Features
The project has been customized with two main modules:
1. **Bill Download** (`/welcome`) - Primary functionality
2. **PDF Tools** (`/pdf/batch-rename`) - PDF batch processing, specifically for invoices

### Routing & Navigation
- Simplified routing system with static routes only
- Basic authentication using cookies with `multipleTabsKey`
- No complex permission system or dynamic routes
- Layout uses standard admin template structure with sidebar navigation

### Development Notes
- Uses `@/` alias for `src/` imports
- TypeScript strict mode disabled for flexibility
- Node.js >=20.19.0 and pnpm >=9 required
- All linting is currently disabled
- Development server runs on port from `VITE_PORT` env variable (default varies)
- Build output organized in `static/` directory with hash-based filenames

### Configuration
- Dynamic config loading from `public/platform-config.json`
- Environment-based builds (development, staging, static)
- Vite configuration includes CDN support, compression, and optimization
- Custom build utilities in `build/` directory for plugins and optimization