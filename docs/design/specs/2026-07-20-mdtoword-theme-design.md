# MDtoWord Theme Refinement Design

## Goal

Replace the partially styled light interface with a complete, polished dark and light theme system that removes the default Qt gray tab-pane artifact.

## Constraints

- Dark mode is the initial theme.
- The user can switch themes from the application footer.
- Theme preference persists between launches through `QSettings`.
- Every visible Qt surface must be themed explicitly; no system-colored tab, list, combo popup, scrollbar, focus-ring, or disabled state may remain.
- Keep the compact file-conversion flow and its existing behavior unchanged.
- Leave all changes uncommitted for user inspection.

## Architecture

A focused `ThemeManager` owns the current theme identifier, stylesheet construction, and persisted preference. The application invokes it at startup and from the theme-toggle action. It applies the Fusion style before assigning a complete QSS stylesheet to the application.

The style definitions use named palette tokens rather than widget-local hardcoded colors. Dark theme uses `#0D1117` for the application background, `#161B22` for surfaces, `#21262D` for borders, `#7C6CFF` for action and selected states, and `#A9B1D6` for secondary text. Light mode mirrors contrast and hierarchy rather than reverting to unstyled platform defaults.

The QSS explicitly styles `QTabWidget::pane` and `QTabBar` to eliminate the screenshot artifact, plus the main window, group boxes, inputs, `QListWidget`, `QAbstractScrollArea`, scrollbars, buttons, combo popup views, spinboxes, progress bars, selected items, focus states, and disabled states.

## Interface

- The existing footer receives a compact theme button with moon/sun icon text and accessible tooltip.
- The primary conversion action uses the shared indigo accent; secondary and destructive actions use subdued theme-specific colors.
- The drop zone, file queue, and output card use layered surfaces and 1px boundaries for clear hierarchy.
- No widget uses the previous pink focus outline or light-only background styling.

## Validation

- Unit-test default dark theme, toggling, persisted preference, and the presence of required QSS selectors including the tab pane.
- Offscreen GUI smoke-test application construction and theme toggle.
- Run the full suite, compile sources, rebuild `dist/MDtoWORD.app`, verify signing, and launch the app bundle.
